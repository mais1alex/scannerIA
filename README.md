# Automação de Vencimento de Documentos de Veículos

Automação em **Google Apps Script** que varre pastas do Google Drive contendo documentos de veículos (PDFs e imagens), extrai dados via IA (Gemini), calcula vencimentos com base em regras configuráveis e calendários estaduais de licenciamento, e envia alertas por e-mail sobre documentos vencidos ou próximos do vencimento.

Todo o estado da automação (banco de arquivos, logs, regras, pendências e calendários) é mantido em uma única planilha Google Sheets criada automaticamente.

## Por que existe

Frotas acumulam dezenas de documentos com regras de validade diferentes: CRLV (controlado por exercício anual + calendário estadual por final de placa), AET, CIV, CIPP, tacógrafo, apólices de seguro, registros ANTT/IBAMA, entre outros. O controle manual falha silenciosamente. Este script transforma uma estrutura de pastas do Drive em um sistema de monitoramento contínuo, com trilha de auditoria e fila de revisão humana para casos ambíguos.

## Como funciona

1. **Varredura com escopo controlado** — apenas as pastas de primeiro nível listadas em `PASTAS_RAIZ_PERMITIDAS` são processadas; pastas em `PASTAS_IGNORADAS` (arquivos mortos, vendidos, ex-sócios etc.) são puladas em qualquer nível.
2. **Detecção de mudanças** — cada arquivo recebe uma assinatura (`id|nome|tamanho|última modificação`). Arquivos já processados e inalterados são pulados, evitando chamadas desnecessárias à IA.
3. **Extração via Gemini** — o arquivo é enviado em base64 para o modelo, que retorna JSON estruturado: tipo de documento, placa, UF, data de emissão, data de vencimento e (para CRLV) o exercício.
4. **Normalização e cálculo** — se o documento não traz vencimento explícito, o script aplica a regra cadastrada para o tipo/UF (ex.: `ANOS_APOS_EMISSAO`). CRLVs são tratados por exercício: o vencimento real é buscado no calendário estadual de licenciamento (UF + ano + final da placa).
5. **Classificação de situação** — `REGULAR`, `URGENTE` (≤ 60 dias), `VENCIDO`, `LICENCIAMENTO_PROXIMO`, `LICENCIAMENTO_ATRASADO`, `VENCIDO_HA_MAIS_DE_UM_ANO`, entre outras.
6. **Pendências e alertas** — casos que a automação não resolve sozinha (sem exercício, sem UF, calendário não cadastrado, erro de leitura) vão para a aba de pendências com ação recomendada. Documentos em alerta disparam e-mail em HTML com tabela e links diretos para os arquivos.

## Estrutura da planilha de controle

A planilha `CONTROLE_IA_DOCUMENTOS_VEICULOS` é criada automaticamente na pasta-mãe com cinco abas:

| Aba | Função |
|---|---|
| `LOG_EXECUCAO` | Trilha de auditoria de cada execução (nível, categoria, mensagem) |
| `ARQUIVOS_PROCESSADOS` | Banco de dados: um registro por arquivo, com assinatura, dados extraídos, vencimento, situação e o JSON bruto da IA |
| `REGRAS_VENCIMENTO` | Regras de cálculo por tipo de documento e UF (editáveis) |
| `PENDENCIAS_REVISAO` | Fila de revisão humana, regenerada a cada execução, com motivo e ação recomendada |
| `CALENDARIO_LICENCIAMENTO` | Calendários estaduais de licenciamento por UF, ano e final de placa |

## Configuração rápida

1. Crie um projeto no [Google Apps Script](https://script.google.com) vinculado à sua conta com acesso ao Drive.
2. Cole o conteúdo de `Code.gs`.
3. Ajuste as constantes no topo do arquivo:
   - `PASTA_MAE_ID` — ID da pasta raiz no Drive
   - `EMAIL_NOTIFICACAO` — destinatário dos alertas
   - `PASTAS_RAIZ_PERMITIDAS` — nomes das pastas de primeiro nível a processar
   - `PASTAS_IGNORADAS` — nomes de pastas a ignorar em qualquer nível
   - `GEMINI_MODEL` — modelo a utilizar
4. Configure a chave da API Gemini (ver seção de segurança abaixo — **não** deixe a chave hardcoded).
5. Execute `testarEscopoDasPastas()` para validar o escopo sem consumir tokens.
6. Execute `verificarVencimentosDeVeiculos()` manualmente para a primeira carga.
7. Crie um acionador (trigger) baseado em tempo para execução recorrente (ex.: diária).

### Limites operacionais

- `MAX_ARQUIVOS_PROCESSAR_POR_EXECUCAO = 25` — protege contra o timeout de 6 minutos do Apps Script. Execuções sucessivas continuam de onde pararam (arquivos já processados são pulados pela assinatura).
- `PAUSA_ENTRE_CHAMADAS_MS = 5000` — respeita rate limits da API Gemini.
- `DIAS_ALERTA = 60` — janela de antecedência para o status `URGENTE` / `LICENCIAMENTO_PROXIMO`.

## Regras de vencimento

Critérios suportados na aba `REGRAS_VENCIMENTO`:

- `USAR_VALIDADE_DOCUMENTO` — usa apenas a validade lida no documento
- `DIAS_APOS_EMISSAO`, `MESES_APOS_EMISSAO`, `ANOS_APOS_EMISSAO` — cálculo relativo à emissão
- `FIM_ANO_EMISSAO`, `FIM_ANO_SEGUINTE` — fim do ano civil
- `CONTROLAR_POR_EXERCICIO` — específico de CRLV (usa calendário estadual)
- `NAO_APLICAVEL` — documentos complementares sem vencimento próprio (ex.: RNC)

As regras genéricas pré-carregadas devem ser validadas conforme o órgão emissor de cada documento.

## Tratamento especial: CRLV e RNC

- **CRLV** nunca tem vencimento inventado a partir da emissão. O script extrai o exercício (via IA, observação ou nome do arquivo), consulta o calendário da UF para `exercício + 1` e o final numérico da placa, e classifica o documento. CRLVs defasados em dois ou mais exercícios são marcados como `VENCIDO_HA_MAIS_DE_UM_ANO` mesmo sem calendário cadastrado.
- **RNC** é documento complementar vinculado à inspeção, sem vencimento próprio — classificado como `NAO_APLICAVEL` e excluído de alertas e pendências.
- A função `reclassificarRegistrosAntigosCRLVeRNC()` reprocessa registros antigos com a lógica atual **sem chamar a IA**, permitindo evoluir as regras sem custo de tokens.

## Funções auxiliares manuais

| Função | Uso |
|---|---|
| `testarEscopoDasPastas()` | Simula o escopo de varredura sem chamar a IA |
| `recalcularTodosOsCRLVsPeloCalendario()` | Recalcula CRLVs pelos calendários cadastrados, sem tokens |
| `limparBancoDeArquivosProcessados()` | Força reprocessamento total (apaga só a aba de banco) |
| `limparLogExecucao()` | Limpa logs antigos |
| `configurarChaveGemini()` / `removerChaveGemini()` | Gerencia a chave nas propriedades do script |

## Segurança e LGPD

- **Nunca** deixe a chave da API no código. Use `configurarChaveGemini()` para armazená-la em `PropertiesService` e garanta que o código a leia de lá (ver observação no código).
- Os documentos processados contêm dados pessoais (placas, nomes de proprietários, possivelmente CPF). A coluna `json_ia` armazena a resposta bruta da IA — restrinja o compartilhamento da planilha de controle ao mínimo necessário.
- O conteúdo dos arquivos é enviado à API do Gemini para análise. Avalie se isso é compatível com a política de tratamento de dados da sua organização antes de processar documentos reais.
- Os calendários de licenciamento pré-carregados devem ser confirmados anualmente nas fontes oficiais dos DETRANs antes da virada de exercício. AP, PA, PR e RN exigem cadastro manual (regras estaduais especiais).

## Limitações conhecidas

- Máximo de 25 arquivos por execução (contornável com trigger recorrente).
- Extração depende da qualidade/legibilidade do documento — casos ilegíveis caem em pendência, não são inventados.
- Apps Script não oferece paralelismo: execuções longas em frotas grandes exigem múltiplos ciclos do trigger.
- A aba de pendências é regenerada a cada execução; anotações manuais feitas nela são perdidas.

## Autoria

Desenvolvido por Alexsandro Silva Borba.

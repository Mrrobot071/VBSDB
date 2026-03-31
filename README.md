# Automação Smart de Extração e Consolidação de Notas Fiscais

Transformando e-mails, anexos e documentos fiscais em um fluxo operacional confiável, rastreável e escalável.

---

## Visão Geral

Este projeto é uma solução de automação construída para resolver um problema real e recorrente no ambiente corporativo: o tratamento manual de e-mails com notas fiscais, boletos, pedidos, anexos e informações de pagamento.

Na prática, ele atua como uma camada de gestão operacional entre o Outlook, os arquivos recebidos e a planilha de controle, automatizando tarefas que normalmente consumiriam tempo, energia e atenção humana em alto volume.

Mais do que uma macro, este projeto representa uma arquitetura pragmática de automação administrativa: ele coleta, filtra, interpreta, organiza e transforma dados dispersos em um fluxo padronizado e utilizável. O resultado é uma operação mais rápida, mais segura e muito menos sujeita a erro humano.

O ecossistema consolidado no arquivo demonstra características de automação robusta, rastreabilidade completa, controle operacional, compliance, integração entre áreas e capacidade de processar grandes volumes de transações, envolvendo pedidos, notas, valores, vencimentos, fornecedores, compradores e fluxos de aprovação.

---

## O Problema que Este Projeto Resolve

Empresas que recebem documentos fiscais por e-mail enfrentam, diariamente, desafios como:

- Localizar mensagens relevantes em caixas compartilhadas;
- Baixar anexos manualmente;
- Ignorar arquivos irrelevantes ou duplicados;
- Identificar pedido, nota fiscal, valor, vencimento e fornecedor em textos despadronizados;
- Consolidar informações em planilhas operacionais;
- Manter rastreabilidade e consistência entre áreas como compras, fiscal e financeiro.

Este projeto resolve esse cenário com uma abordagem extremamente eficiente:

1. Vasculha a pasta correta no Outlook;
2. Baixa anexos automaticamente;
3. Ignora imagens e duplicidades;
4. Mantém o nome original do arquivo com segurança;
5. Extrai dados estruturados dos e-mails e PDFs;
6. Organiza tudo em planilhas prontas para operação;
7. Permite consulta cruzada com a base IARA.

---

## Por que este projeto é importante?

O diferencial deste projeto não está apenas em “automatizar uma tarefa”.

Ele está em algo mais raro: traduzir um processo operacional caótico em uma engenharia confiável de execução.

### 1. Ele entende o fluxo real do negócio

O projeto não foi feito em laboratório. Ele nasce do mundo real — com exceções, ruído, nomes inconsistentes, anexos duplicados, e-mails encaminhados, padrões irregulares e necessidade de resposta rápida.

### 2. Ele automatiza com critério, não com ingenuidade

Não é uma automação “cega”.

Há regras de exclusão, validação de extensões, sanitização de nomes, verificação de duplicidade, busca por pasta específica, filtros por mensagens não lidas e tratamento de erro.

### 3. Ele conecta operação e inteligência

A solução não apenas baixa arquivos. Ela estrutura conhecimento operacional:

- Pedido;
- NF;
- Valor;
- Vencimento;
- Forma de pagamento;
- Comprador;
- Fornecedor;
- Observações;
- Rastreio da origem.

Esse tipo de organização cria um ativo valioso: dados utilizáveis para controle, conferência e decisão.

### 4. Ele escala

O material consolidado mostra um fluxo capaz de lidar com centenas de transações, múltiplos fornecedores, diferentes compradores, faixas amplas de valores e operações envolvendo várias áreas da empresa.

### 5. Ele reduz atrito entre áreas

Compras, financeiro, fiscal, jurídico, eventos, comunicação, TI e operação passam a trabalhar sobre um fluxo mais padronizado, rastreável e previsível. O próprio conjunto do projeto evidencia essa integração intersetorial e a documentação detalhada de cada transação.

### 6. Ele tem mentalidade de produto

Mesmo sendo desenvolvido em VBA/Excel/Outlook, o projeto possui características de produto interno:

- Resolve dor real;
- Aumenta produtividade;
- Reduz risco operacional;
- Melhora compliance;
- Padroniza execução;
- Cria histórico.

---

## Arquitetura da Solução

A solução pode ser entendida em três camadas principais:

### 1. Captura

Responsável por acessar a caixa de entrada no Outlook e localizar a pasta correta de trabalho.

No trecho apresentado, a macro procura a estrutura:

`NFE > 01.NOTAS`

e processa apenas os e-mails relevantes.

### 2. Tratamento

Essa camada aplica regras de negócio para garantir qualidade operacional:

- Ignora anexos de imagem (jpg, png, gif, bmp, webp etc.);
- Evita salvar arquivos duplicados;
- Sanitiza nomes de arquivos para evitar caracteres inválidos;
- Cria a pasta de destino automaticamente;
- Preserva o nome original dos anexos;
- Pode marcar e-mails como lidos após processamento.

Essa combinação mostra um cuidado raro com resiliência de execução — algo essencial em automações corporativas.

### 3. Estruturação e Consolidação

Além do download de anexos, o projeto expande o fluxo para:

- Extrair campos relevantes de e-mails e PDFs;
- Identificar pedido e nota fiscal;
- Capturar valor e vencimento;
- Reconhecer fornecedor e comprador;
- Preencher planilhas estruturadas;
- Cruzar dados com a planilha IARA por meio de busca automatizada.

O conjunto consolidado do arquivo demonstra rastreabilidade total, com cada transação associada a pedido, NF, valor, vencimento, fornecedor, comprador, observações, timestamps e histórico operacional.

---

## Principais Funcionalidades

### Download inteligente de anexos

- Leitura automática de e-mails na pasta operacional;
- Salvamento em diretório padronizado;
- Exclusão de anexos irrelevantes;
- Prevenção de duplicidade.

### Higienização de arquivos

- Remoção de caracteres inválidos;
- Padronização segura do nome do arquivo;
- Compatibilidade com Windows e OneDrive.

### Leitura de informações operacionais

- Identificação de pedido;
- Identificação de nota fiscal;
- Leitura de valor;
- Leitura de vencimento;
- Captura de forma de pagamento;
- Captura de comprador e fornecedor.

### Consolidação em planilha

- Geração de base estruturada;
- Preenchimento de colunas operacionais;
- Suporte a conferência e programação de pagamento;
- Integração com referência externa (IARA).

### Robustez operacional

- Tratamento de erros;
- Filtros por e-mail não lido;
- Suporte a grande volume;
- Controle de existência de arquivo;
- Estrutura preparada para expansão.

---

## Diferenciais Técnicos

Este projeto se destaca por várias decisões de implementação muito inteligentes:

### Restrict no Outlook

O uso de filtro para mensagens não lidas melhora performance e evita processamento desnecessário.

### Scripting.Dictionary para controle de duplicidade

Uma escolha elegante para impedir regravações no mesmo ciclo de execução.

### Sanitização de nomes com SafeFileComponent

Evita falhas de sistema de arquivos e reduz risco de inconsistência.

### Criação recursiva de diretórios

A automação não depende de preparo manual do ambiente.

### Separação entre sub principal e funções auxiliares

Isso melhora manutenção, clareza e evolução futura.

### Estratégia orientada a negócio

A macro não apenas “executa código”; ela reflete regras operacionais reais.

---

## Impacto Operacional

Com base no comportamento da solução e no conteúdo consolidado no workbook, os ganhos são claros:

- Mais velocidade no tratamento de documentos;
- Menos retrabalho manual;
- Menos erro humano;
- Mais padronização;
- Mais rastreabilidade;
- Mais segurança da informação;
- Mais previsibilidade no fluxo administrativo e financeiro.

O material também evidencia preocupações com compliance, confidencialidade, controle de acesso, mitigação de risco e padronização de comunicação, o que eleva o projeto de simples automação para ferramenta crítica de gestão operacional.

---

## Tecnologias Utilizadas

- VBA (Visual Basic for Applications)
- Microsoft Outlook Object Model
- Microsoft Excel
- Power Query / M
- FileSystemObject
- OneDrive / ambiente corporativo Windows

---

## Estrutura Conceitual do Fluxo

Outlook (`NFE > 01.NOTAS`)  
↓  
Leitura de e-mails não lidos  
↓  
Filtragem de anexos  
↓  
Remoção de imagens e duplicados  
↓  
Salvamento seguro dos arquivos  
↓  
Extração de dados do corpo do e-mail / PDFs  
↓  
Estruturação em planilha  
↓  
Conferência / programação / integração com IARA

---

## Casos de Uso

- Contas a pagar;
- Conferência de notas fiscais;
- Abertura de chamados operacionais;
- Programação de pagamentos;
- Auditoria e rastreabilidade;
- Apoio a compras e fiscal;
- Organização de anexos corporativos;
- Consolidação de dados de fornecedores.

---

## O que torna esse projeto especial

Há muitos scripts que “funcionam”.

Poucos realmente entendem a operação.

Este projeto é especial porque combina:

- Visão de processo;
- Domínio prático da rotina;
- Cuidado com exceções;
- Senso de organização;
- Automação com inteligência;
- Foco em confiabilidade.

Em outras palavras:

Ele não é apenas código.

Ele é engenharia aplicada ao caos administrativo.

E isso é exatamente o que torna o projeto genial.

---

## Possíveis Evoluções Futuras

- Interface para parametrização sem editar VBA;
- Logs mais detalhados em aba dedicada;
- Versionamento automático de arquivos;
- Dashboard com métricas de processamento;
- Integração com banco de dados;
- Classificação automática por tipo de documento;
- Validações adicionais de CNPJ/NF;
- Alertas de inconsistência antes da consolidação;
- Empacotamento como solução interna escalável.

---

## Conclusão

Este projeto demonstra como ferramentas aparentemente simples — Outlook, Excel, VBA e Power Query — podem ser transformadas em uma solução de alto valor quando combinadas com visão operacional, disciplina lógica e foco em eficiência.

Ele automatiza, organiza, valida, consolida e dá rastreabilidade a um fluxo crítico.

E faz isso de forma elegante, útil e extremamente inteligente.

Essa é a genialidade do projeto:

Pegar um processo que normalmente depende de esforço manual repetitivo e convertê-lo em uma máquina confiável de execução administrativa.

Na essência, o projeto pode ser considerado um ETL, com uma observação importante: ele é um ETL orientado a documentos (e-mails/PDFs/anexos) e com perfil de automação/RPA, porque a “fonte” não é um banco de dados, e sim Outlook + arquivos.

A forma mais precisa de descrever seria:

✅ “Pipeline ETL (document-driven) / ETL operacional”

ou

✅ “Automação + ETL para documentos fiscais (Outlook → Excel/IARA)”

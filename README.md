# Análise de Medidas Disciplinares

Dashboard local em HTML para uso do RH, voltado à leitura de planilhas Excel e geração automática de KPIs de medidas disciplinares.

## Visão geral

Este projeto foi construído como uma aplicação de arquivo único, pensada para rodar localmente no navegador, sem backend e sem envio de dados para servidores da aplicação.

O objetivo é permitir que a gestão de RH:

- importe uma planilha de medidas disciplinares;
- visualize indicadores executivos rapidamente;
- investigue ocorrências por tipo, departamento, centro de custo e colaborador;
- exporte um snapshot do dashboard em HTML para compartilhamento interno.

## Arquitetura do projeto

Estrutura atual do repositório:

- `Analise de Medidas Disciplinares GMINZ v3.html`: aplicação completa em HTML, CSS e JavaScript.
- `README.md`: documentação funcional e técnica do projeto.

### Características principais

- aplicação 100% client-side;
- processamento local no navegador;
- interface única com tela de importação e dashboard;
- leitura de arquivos `.xlsx` e `.xls`;
- dados de demonstração embutidos;
- exportação do dashboard como um novo HTML independente.

## Como usar

1. Abra o arquivo `Analise de Medidas Disciplinares GMINZ v3.html` em um navegador moderno.
2. Arraste uma planilha Excel para a área de upload ou clique para selecionar o arquivo.
3. Aguarde o processamento local dos dados.
4. Navegue pelos KPIs, listas e modais de detalhamento.
5. Se quiser compartilhar o resultado, clique em **Exportar dashboard**.

### Navegadores recomendados

- Google Chrome
- Microsoft Edge

O sistema deve funcionar em navegadores modernos compatíveis com:

- `FileReader`
- `Blob`
- `URL.createObjectURL`
- `Date`
- `Array.map/filter/reduce`

## Dependências externas

O projeto é local, mas atualmente depende de recursos carregados da internet:

- biblioteca SheetJS via CDN para leitura de Excel;
- fontes Google Fonts.

Isso significa que:

- os dados da planilha são processados localmente no navegador;
- porém a interface pode precisar de acesso à internet para carregar a biblioteca externa e as fontes, especialmente na primeira abertura.

Se houver necessidade de uso totalmente offline, o ideal é versionar localmente a biblioteca de Excel e substituir as fontes externas por assets locais ou fontes do sistema.

## Fluxo funcional

### 1. Importação da base

A aplicação lê apenas a **primeira aba** do arquivo Excel.

Na importação, ela:

- converte a planilha para JSON;
- normaliza nomes de colunas removendo espaços extras;
- descarta linhas sem colaborador identificado;
- transforma e enriquece os registros para cálculo dos indicadores.

### 2. Tratamento dos dados

Para cada linha válida, o sistema monta um registro com campos padronizados, incluindo:

- empresa;
- data;
- grupo/departamento;
- classe/centro de custo;
- funcionário;
- sequência;
- tipo;
- matrícula;
- anotações.

Também cria campos derivados, como:

- descrição do tipo (`TipoDesc`);
- peso de severidade (`Peso`);
- flag de reincidência (`Reincidente`);
- agrupamento por ano e mês (`AnoMes`);
- categoria temática inferida a partir da anotação (`Categoria`).

### 3. Exibição do dashboard

Após o processamento, a tela de upload é ocultada e o dashboard é exibido com:

- cards de KPI;
- distribuição por tipo;
- insights executivos automáticos;
- evolução mensal;
- ranking por departamento;
- ranking por centro de custo;
- tabela paginada de colaboradores;
- modais com detalhamento analítico.

## Formato esperado da planilha

A aplicação foi desenhada para reconhecer principalmente colunas com os seguintes nomes:

- `Empresa`
- `Data`
- `Grupo`
- `Classe`
- `Funcionario`
- `Sequencia`
- `Tipo`
- `Matricula`
- `Anotacoes`

### Observações importantes

- os nomes são tratados com tolerância limitada a maiúsculas/minúsculas e espaços nas chaves;
- a lógica assume esses nomes em português, sem uma camada completa de mapeamento flexível;
- linhas sem `Funcionario` são descartadas;
- a primeira aba do Excel é a única considerada;
- se a planilha estiver vazia ou sem colunas reconhecíveis, o sistema exibe erro.

## Regras de negócio implementadas

### Mapeamento de tipos de medida

O campo `Tipo` é convertido para uma descrição amigável:

- `2` → Advertência escrita
- `3` → Advertência verbal
- `4` → Suspensão

Se o código não estiver mapeado, a interface exibe `Tipo não mapeado`.

### Peso de severidade

O sistema aplica uma escala simples para análise de gravidade:

- `Advertência verbal` = peso 1
- `Advertência escrita` = peso 2
- `Suspensão` = peso 4

Esse peso é usado para identificar o tipo mais grave por colaborador.

### Reincidência

Um registro é tratado como reincidência quando:

- `Sequencia > 1`

A contagem total de reincidências e a marcação por colaborador derivam dessa regra.

### Classificação automática das anotações

O campo `Anotacoes` é analisado por regras de palavras-chave para enquadrar a ocorrência em categorias temáticas.

Categorias atualmente implementadas:

- Segurança operacional
- Conduta e convivência
- Assiduidade e jornada
- Procedimento e comunicação
- Patrimônio e instalações
- Outros
- Não informado

Essa classificação é heurística, baseada em regex e normalização textual. Ou seja, ajuda na leitura gerencial, mas não substitui classificação humana formal.

## KPIs e análises geradas

### Indicadores gerais

O dashboard monta os seguintes indicadores principais:

- total de medidas;
- colaboradores impactados;
- reincidências (`Sequencia > 1`);
- tipo mais frequente;
- departamento com mais medidas;
- centro de custo com mais medidas;
- período analisado.

### Distribuições

São calculadas distribuições por:

- tipo de medida;
- mês;
- categoria temática;
- departamento (`Grupo`);
- centro de custo (`Classe`).

### Análises por colaborador

A visão de colaboradores inclui:

- total de ocorrências por pessoa;
- tipo mais grave recebido;
- última ocorrência;
- presença de reincidência;
- detalhamento por tipo;
- busca por nome, matrícula, grupo ou classe;
- paginação da tabela.

### Insights executivos automáticos

O sistema escreve mensagens interpretativas com base nos dados, como:

- tipo predominante;
- mês de pico;
- percentual de reincidência;
- departamento com maior concentração;
- centro de custo com maior concentração.

## Interações disponíveis

### Tela inicial

- upload por clique;
- upload por arrastar e soltar;
- botão de demonstração com dados fictícios.

### Dashboard

- botão **Nova importação** para reiniciar o fluxo;
- botão **Exportar dashboard** para baixar um HTML estático com os dados já embutidos.

### Drill-down analítico

É possível abrir modais a partir de:

- tipos de medida;
- departamentos;
- centros de custo;
- histórico individual do colaborador.

Esses modais permitem investigação detalhada sem sair da tela principal.

## Exportação

Ao clicar em **Exportar dashboard**, o sistema gera um novo arquivo HTML com:

- os dados atuais embutidos em JSON;
- a tela de upload oculta;
- a renderização direta do dashboard ao abrir o arquivo exportado;
- botões de importação/exportação removidos do snapshot.

### Uso prático da exportação

Esse recurso é útil para:

- enviar uma versão estática do painel para liderança;
- arquivar um fechamento mensal;
- manter um histórico de snapshots por período.

## Privacidade e segurança

### O que o projeto faz bem

- processa a planilha no navegador local;
- não possui backend próprio;
- não envia os registros para uma API da aplicação.

### Atenções importantes

Como o HTML carrega bibliotecas e fontes externas, existe dependência de terceiros para montagem da página. Embora os dados da planilha sejam processados localmente, para um ambiente corporativo com maior exigência de segurança é recomendável:

- hospedar ou embutir localmente a biblioteca de leitura de Excel;
- remover dependências de fontes externas;
- validar o uso em ambiente offline controlado;
- revisar política interna de tratamento de dados pessoais.

## Limitações atuais

Pelos comportamentos implementados hoje, os principais limites são:

- leitura apenas da primeira aba do Excel;
- ausência de filtros dinâmicos por período, empresa, departamento ou tipo;
- ausência de persistência local entre sessões;
- ausência de autenticação e controle de acesso;
- classificação temática baseada em regras simples de texto;
- dependência de internet para recursos CDN/fonts;
- ausência de testes automatizados;
- ausência de separação entre camada visual e lógica de negócio.

## Manutenção e evolução recomendada

Se o projeto continuar crescendo, estes são os próximos passos mais valiosos:

### Curto prazo

- criar um dicionário de colunas aceitas com aliases;
- documentar um modelo oficial de planilha para o RH;
- melhorar mensagens de erro de importação;
- incluir filtros por período, tipo, grupo e classe;
- mostrar quantidade de registros inválidos descartados.

### Médio prazo

- separar o JavaScript em módulos;
- externalizar CSS e regras de negócio;
- adicionar testes para o cálculo de KPIs;
- salvar preferências/filtros localmente;
- incluir exportação para PDF ou imagem.

### Segurança e operação

- empacotar dependências para uso totalmente offline;
- padronizar distribuição em pasta única para o RH;
- controlar versões dos snapshots exportados;
- revisar anonimização quando necessário para apresentações gerenciais.

## Exemplo de base mínima

Uma base funcional deve conter, no mínimo, dados equivalentes a:

| Empresa | Data | Grupo | Classe | Funcionario | Sequencia | Tipo | Matricula | Anotacoes |
| --- | --- | --- | --- | --- | ---: | ---: | --- | --- |
| Empresa X | 2026-03-01 | Operação Mina | CC-101 | JOÃO SILVA | 1 | 2 | 1234 | Descumprimento de procedimento operacional |
| Empresa X | 2026-03-15 | Operação Mina | CC-101 | JOÃO SILVA | 2 | 3 | 1234 | Reincidência em conduta inadequada |
| Empresa X | 2026-03-20 | Manutenção | CC-205 | MARIA SOUZA | 1 | 4 | 5678 | Falha grave com risco de segurança |

## Público-alvo

Este projeto atende especialmente:

- gestores de RH;
- analistas de administração de pessoal;
- liderança operacional que precise acompanhar disciplina por área;
- equipes que precisam apresentar KPIs rapidamente sem depender de sistemas corporativos complexos.

## Resumo executivo do estado atual

Hoje o projeto já entrega bem o objetivo principal: transformar uma planilha operacional em um painel visual executivo, local e fácil de compartilhar.

Os maiores diferenciais atuais são:

- simplicidade de distribuição;
- velocidade de uso;
- leitura gerencial clara;
- detalhamento por colaborador, departamento, classe e tipo;
- exportação em HTML estático.

Os principais riscos de continuidade estão em:

- manutenção de regras de negócio dentro de um único arquivo grande;
- dependências externas para um cenário que idealmente deveria ser offline;
- fragilidade a variações de layout/nome de colunas na planilha.

## Arquivo principal

Para operar ou evoluir o projeto, o arquivo central é:

- `Analise de Medidas Disciplinares GMINZ v3.html`

Se quiser expandir o projeto, o próximo passo mais recomendável é separar:

- interface;
- parsing da planilha;
- cálculo de indicadores;
- exportação;
- regras de classificação.

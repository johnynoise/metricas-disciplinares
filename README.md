# KPIs de Medidas Disciplinares (RH)

Este repositório possui agora **duas formas de uso**:

1. **Gerador por script** (lote/automação);
2. **Sistema local com interface executiva** para importar Excel, filtrar e gerar KPIs.

## O que o sistema entrega

- leitura automática do Excel de medidas disciplinares;
- tratamento da base e padronização das colunas;
- KPIs executivos para RH;
- classificação automática dos motivos por categoria;
- exportação em `CSV`, `JSON`, `XLSX` e `HTML`;
- painel visual local com filtros, gráficos e download em `.zip`.

## KPIs incluídos

- total de medidas disciplinares;
- total de colaboradores impactados;
- medidas com `Sequencia > 1`;
- taxa de medidas com repetição;
- distribuição por tipo de medida;
- severidade média ponderada;
- evolução mensal;
- categorias mais recorrentes nas anotações;
- colaboradores com maior severidade;
- palavras mais frequentes nas descrições.

## Assunção usada para o campo `Tipo`

O script assume o seguinte mapeamento padrão:

- `002 = Advertência disciplinar`
- `003 = Advertência verbal`
- `004 = Suspensão`

Na coluna `Sequencia`, o valor representa a ordem da medida aplicada ao colaborador. Assim, `Sequencia > 1` indica que ele já recebeu mais de uma medida.

Se o seu RH usar outra convenção, basta ajustar o dicionário `TIPO_MEDIDA_MAP` no arquivo `gerar_kpis_medidas_disciplinares.py`.

## Como executar

Instalar dependências:

```bash
python -m pip install -r requirements.txt
```

### 1) Sistema local com interface (recomendado para gestão)

```bash
streamlit run app_kpis_medidas_disciplinares.py
```

No painel:

- faça upload do arquivo Excel;
- aplique filtros por empresa, período e tipo;
- acompanhe os KPIs, gráficos e insights executivos;
- baixe um pacote `.zip` com os artefatos;
- opcionalmente salve os arquivos diretamente em uma pasta local.

### 2) Geração por script (modo tradicional)

Gerar os arquivos de KPI:

```bash
python gerar_kpis_medidas_disciplinares.py
```

Ou informando outros caminhos:

```bash
python gerar_kpis_medidas_disciplinares.py --input outro_arquivo.xlsx --output-dir resultado_rh
```

## Arquivos gerados

Na pasta de saída serão criados:

- `medidas_tratadas.csv`
- `kpis_resumo.json`
- `kpis_medidas_disciplinares.xlsx`
- `dashboard_kpis_medidas_disciplinares.html`

## Observação

A classificação por categoria é feita por palavras-chave nas anotações. Ela já ajuda bastante na análise gerencial, mas pode ser refinada depois com regras específicas da sua operação e do seu regulamento interno.

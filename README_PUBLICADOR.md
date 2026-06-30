# Publicador de Escalas → Lovable (Supabase)

`publicar_lovable.py` grava uma escala **completa e consistente** direto no banco do
Lovable, num passo só — substituindo as duas importações da aba "Visualizar Escalas",
que são incompletas e divergem entre si (uma não calcula carga horária, a outra não
monta o plano de blocos).

## O que ele grava (de forma coerente, numa transação)

1. `rodizios_escala` — datas, nº de semanas, `opcao_total_sg`
2. `specialty_block_config` — blocos com nomes
3. `block_student_distribution` — bloco → serviço (catálogo mestre) + horários
4. `block_week_assignments` — **subgrupo × semana → bloco** (fonte do painel e do Excel)
5. `students.subgrupo` — divisão dos subgrupos
6. `semanas_rodizio` — semanas com datas
7. `escalas_diarias` — serviço por período + **carga horária calculada** + área verde

## Entradas (saída do gerador)

- `template_*.xlsx` (abas `_Blocos`, `Distribuição`, `_Serviços`, `Semana Padrão`, `_Subgrupos`)
- `importar_lovable_*.xlsx` (aba `Correções <cod>`)
- `definicoes_*.json`

## Como usar

### 1. Configurar o segredo (uma vez)

Pegue a connection string no Supabase: **Project Settings → Database → Connection string
(URI) → aba "Session pooler"**, e substitua `[YOUR-PASSWORD]` pela senha do banco.

Crie um arquivo `.env` (já está no `.gitignore`, nunca vai pro git):

```
SUPABASE_DB_URL=postgresql://postgres.xxxx:SUASENHA@aws-0-sa-east-1.pooler.supabase.com:5432/postgres
```

No Windows, alternativamente, na hora de rodar:
```
set SUPABASE_DB_URL=postgresql://postgres.xxxx:SUASENHA@...:5432/postgres
```

### 2. Instalar dependências

```
pip install -r requirements.txt
```

### 3. Validar (não grava nada)

```
python publicar_lovable.py --template T.xlsx --importar I.xlsx --definicoes D.json \
    --codigo R1-GO-T6 --dry-run
```

### 4. Publicar

```
python publicar_lovable.py --template T.xlsx --importar I.xlsx --definicoes D.json \
    --codigo R1-GO-T6 --opcao-sg 8
```

A operação é **idempotente por rodízio**: rodar de novo substitui os dados daquele
rodízio (não duplica). Em caso de erro, faz **rollback** (não grava nada pela metade).

## Validação

A lógica foi validada reproduzindo exatamente a escala R1-GO-T6 já correta no banco:
CH = 4763h, 702 linhas diárias, 168 atribuições de bloco, rotação diagonal por subgrupo.

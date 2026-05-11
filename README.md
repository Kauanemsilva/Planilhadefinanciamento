# 🏠 Planilhas de Imóvel Financiado

<img width="1985" height="600" alt="ChatGPT Image 4 de mai  de 2026, 09_28_00" src="https://github.com/user-attachments/assets/24298d8d-a200-4968-b8b4-d57c3b1e8820" />


Suite completa de planilhas Excel para controle de imóvel financiado na planta — do sinal até a quitação total do financiamento bancário.

<p align="center">
  <img src="https://img.shields.io/badge/Excel-.xlsx-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" />
  <img src="https://img.shields.io/badge/SAC%20%26%20PRICE-360%20parcelas-blue?style=for-the-badge" />
  <img src="https://img.shields.io/badge/F%C3%B3rmulas-Autom%C3%A1ticas-success?style=for-the-badge" />
  <img src="https://img.shields.io/badge/Zero-Erros-brightgreen?style=for-the-badge" />
</p>

---

# 📦 Arquivos do Projeto

## 📋 Imovel_Financiado_PREMIUM.xlsx

Planilha completa para gestão pré-chaves com 4 abas integradas:

* 🏠 Painel Geral
* 🏦 Entrada
* ⚙️ Evolução da Obra
* ✅ Checklist de Documentos

---

## 💳 Financiamento_SAC_PRICE.xlsx

Tabela de amortização standalone com suporte aos sistemas:

* SAC
* PRICE
* Até 360 parcelas
* Cálculos automáticos

### Recursos

* 💳 Controle de Financiamento
* 📉 Gráfico de Saldo Devedor
* 📊 Gráfico Amortização vs Juros

---

## 📖 Manual_Financiamento.docx

Manual completo em Word contendo:

* Configuração inicial
* Diferenças entre SAC e PRICE
* Perguntas frequentes
* Guia passo a passo

---

# 🚀 Começando em 4 Passos

## 1️⃣ Preencha os 6 campos amarelos

Configure as células:

```txt
C6 até C11
```

Informações necessárias:

* Saldo devedor
* Taxa de juros anual
* Número de parcelas
* Seguro mensal
* Sistema (SAC ou PRICE)
* Mês inicial

---

## 2️⃣ Confirme o sistema na célula E11

A célula exibirá automaticamente:

```txt
✅ SAC ativo
```

ou

```txt
✅ PRICE ativo
```

Se aparecer aviso de erro, revise o valor informado em `C10`.

---

## 3️⃣ Registre os pagamentos mensalmente

Preencha:

| Coluna | Informação        |
| ------ | ----------------- |
| K      | Valor pago        |
| L      | Data do pagamento |

A situação da parcela será atualizada automaticamente.

---

## 4️⃣ Acompanhe os KPIs e gráficos

Todos os indicadores atualizam em tempo real:

* Barra de progresso
* Cards KPI
* Gráficos financeiros
* Evolução do saldo

---

# 📊 Colunas da Tabela de Financiamento

| Col. | Campo                   | Tipo       | Descrição                        |
| ---- | ----------------------- | ---------- | -------------------------------- |
| C    | 📅 Mês                  | automático | Calculado pelo mês inicial       |
| E    | 💰 Saldo Devedor        | automático | Redução progressiva do saldo     |
| F    | 💳 Parcela              | automático | SAC reduz / PRICE fixa           |
| G    | 🔨 Amortização          | automático | Parte que reduz o saldo          |
| H    | ⚡ Amortização Adicional | manual     | Pagamento extra opcional         |
| I    | 📊 Juros                | automático | Juros calculados automaticamente |
| J    | 🛡️ Seguro              | automático | Valor definido na configuração   |
| K    | 💵 Valor Pago           | manual     | Valor efetivamente pago          |
| L    | 📆 Data Pgto            | manual     | Data do pagamento                |
| M    | 🚦 Situação             | automático | Status automático da parcela     |

---

# 🚦 Situações Automáticas

| Status     | Regra                          |
| ---------- | ------------------------------ |
| ✅ PAGO     | Valor pago ≥ parcela           |
| ⚠️ PARCIAL | Valor pago menor que a parcela |
| ⏳ PENDENTE | Sem pagamento registrado       |
| 🏁 QUITADO | Saldo devedor zerado           |

---

# 📐 SAC vs 💳 PRICE

## 📐 Sistema SAC

* Amortização fixa
* Parcela reduz mensalmente
* Menor custo total de juros
* Parcela inicial mais alta

---

## 💳 Sistema PRICE

* Parcela fixa
* Amortização cresce ao longo do tempo
* Maior incidência de juros
* Melhor previsibilidade financeira

---

# 🔄 Troca de Sistema

Para alterar entre SAC e PRICE:

1. Vá até a célula:

```txt
C10
```

2. Digite:

```txt
SAC
```

ou

```txt
PRICE
```

⚠️ Sem acentos.

Toda a tabela será recalculada automaticamente mantendo o histórico de pagamentos.

---

# 🧩 Abas da Planilha Premium

## 🏠 Painel Geral

Resumo financeiro completo do imóvel:

* Dados do empreendimento
* Valores pagos
* Indicadores financeiros
* Resumo geral

---

## 🏦 Entrada

Controle da fase pré-chaves:

* Sinal
* Parcelas da entrada
* Status automático

---

## ⚙️ Evolução da Obra

Acompanhamento mensal da construção:

* Percentual acumulado
* Pagamentos da obra
* Gráficos automáticos

---

## ✅ Checklist Docs

Controle documental completo:

* Documentos pessoais
* Documentos do imóvel
* Marcos do financiamento
* Pós-chaves

---

# 🎯 Recursos Principais

✅ Compatível com SAC e PRICE
✅ Até 360 parcelas
✅ Fórmulas automáticas
✅ Dashboard financeiro
✅ Controle de amortização extra
✅ Indicadores em tempo real
✅ Gráficos automáticos
✅ Controle pré e pós-chaves
✅ Manual completo incluso

---

# 🛠️ Tecnologias Utilizadas

* Microsoft Excel (.xlsx)
* Fórmulas avançadas
* Validação de dados
* Formatação condicional
* KPIs automáticos
* Dashboards financeiros

---

# 📄 Licença

Este projeto é destinado para uso educacional e organizacional.

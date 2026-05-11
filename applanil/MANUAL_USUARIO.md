# 📘 Manual de Usuário - Tabela de Amortização de Financiamento Imobiliário

## 📋 Sumário

1. [Visão Geral](#visão-geral)
2. [Como Começar](#como-começar)
3. [Seção de Configurações](#seção-de-configurações)
4. [Preenchimento da Tabela](#preenchimento-da-tabela)
5. [Interpretando os Resultados](#interpretando-os-resultados)
6. [Dicas Avançadas](#dicas-avançadas)
7. [Perguntas Frequentes](#perguntas-frequentes)

---

## 🏠 Visão Geral

A **Tabela de Amortização de Financiamento Imobiliário** é uma planilha interativa que permite:

✅ **Calcular parcelas** automaticamente (sistemas PRICE e SAC)  
✅ **Acompanhar pagamentos** mês a mês  
✅ **Controlar juros e amortizações adicionais**  
✅ **Visualizar saldo devedor** em tempo real  
✅ **Gerar relatórios** de quitação e progresso

### Recursos Principais

- **360 meses** de previsão (30 anos)
- **Dois sistemas** de cálculo: PRICE (parcela fixa) e SAC (parcela decrescente)
- **Amortização adicional** para acelerar quitação
- **Seguro mensal** configurável
- **Formatação visual** com cores para fácil interpretação
- **Totalizadores** automáticos

---

## 🚀 Como Começar

### Passo 1: Abrir a Planilha

1. Execute o arquivo `create_excel_imovel.py`
2. Será gerado o arquivo `Imovel_Financiado_PERSONALIZADO.xlsx`
3. Abra o arquivo no Excel, Google Sheets ou LibreOffice Calc

### Passo 2: Entender a Estrutura

A planilha possui três áreas principais:

- **Zona Roxa (linhas 3-7):** Configurações do financiamento
- **Zona Verde (linhas 10-11):** Totalizadores
- **Zona Azul (linhas 13+):** Tabela de dados mensal

---

## ⚙️ Seção de Configurações

### Localização

As configurações estão na **seção superior roxa** (linhas 4-7)

### Campos Obrigatórios (lado esquerdo)

#### 💰 Valor Financiado (R$)

- **Exemplo:** `500.000,00`
- **O que é:** O valor total que você pediu emprestado
- **Como preencher:** Digite o valor em reais, com ou sem formatação
- ⚠️ **Importante:** Este valor é a base para todos os cálculos

#### 📅 Mês/Ano de Início

- **Exemplo:** `Jan/2025` ou `January/2025`
- **O que é:** Quando começa seu financiamento
- **Como preencher:** Digite no formato `Mês/Ano` (português) ou `Mês/Ano` (inglês)
- 💡 **Dica:** A planilha calculará automaticamente os próximos 360 meses

#### ⏳ Prazo (meses)

- **Exemplo:** `360`
- **O que é:** Quantos meses levará para pagar o financiamento
- **Como preencher:** Digite apenas números
- 📊 **Valores típicos:**
  - 120 meses = 10 anos
  - 180 meses = 15 anos
  - 240 meses = 20 anos
  - 360 meses = 30 anos

#### 📈 Taxa de Juros (% a.m.)

- **Exemplo:** `0,75%` ou `0.75%`
- **O que é:** Percentual de juros mensal
- **Como preencher:** Digite como percentual (ex: `0,75%` para 0,75% ao mês)
- 💡 **Dica:** A taxa não muda ao longo do tempo nesta planilha (taxa fixa)

### Campos Opcionais (lado direito)

#### 🏦 Sistema

- **Opções:** `SAC` ou `PRICE`
- **O que é:** Método de cálculo das parcelas
- **Como escolher:**
  - **PRICE:** Parcela fixa o tempo todo (valor igual todo mês)
  - **SAC:** Parcela diminui ao longo do tempo (mais amortização, menos juros)

| Aspecto              | PRICE                | SAC              |
| -------------------- | -------------------- | ---------------- |
| **Parcela**          | Fixa                 | Diminui          |
| **Primeira parcela** | Média                | Maior            |
| **Últimas parcelas** | Média                | Menor            |
| **Total de juros**   | Mais alto            | Mais baixo       |
| **Melhor para**      | Orçamento previsível | Quer quitar logo |

#### 🛡️ Seguro Mensal (R$)

- **Exemplo:** `250,00`
- **O que é:** Seguro residencial mensal (opcional)
- **Como preencher:** Digite o valor em reais
- ℹ️ **Se deixar em branco:** Será considerado como R$ 0,00
- 💡 **Dica:** Verifique com seu banco o valor real

#### 📊 Amortização Adicional Padrão

- **Exemplo:** `1.000,00`
- **O que é:** Valor extra que você quer pagar por padrão em cada mês
- **Como preencher:** Digite em reais
- 💡 **Dica:** Deixe em branco se não souber ainda; pode preencher mês a mês depois

### Informações Calculadas Automaticamente

#### 📋 Parcela PRICE

- Mostra o valor de cada parcela no sistema PRICE
- **Cálculo:** Feito automaticamente com base nos valores acima

#### 📋 Parcela SAC (referência)

- Mostra o valor da primeira parcela no sistema SAC
- **Nota:** Diminui ao longo do tempo no SAC

#### 💰 Total a Pagar

- Soma de todas as parcelas (principal + juros + seguros)
- **Exemplo:** Se financiar R$ 500 mil a 0,75% a.m. por 360 meses, pode pagar R$ 1,2 milhão

#### 💸 Total de Juros

- Quanto você pagará em juros
- **Cálculo:** Total a Pagar - Valor Financiado

---

## 📊 Preenchimento da Tabela

### Estrutura das Colunas

#### Coluna A (Decorativa)

- Apenas decoração visual (cores azuis alternadas)

#### Coluna B - # (Número da Parcela)

- Preenchida automaticamente
- Vai de 1 a 360

#### Coluna C - 📅 MÊS / ANO

- Preenchida automaticamente
- Exemplo: `Jan/2025`, `Fev/2025`, etc.

#### Coluna D - 📆 VENCIMENTO

- Data de vencimento (dia 5 de cada mês)
- Preenchida automaticamente
- Formato: `DD/MM/YYYY`

#### Coluna E - 🏦 SIST. (Sistema)

- Mostra qual sistema está usando (SAC ou PRICE)
- Preenchida automaticamente com base em sua escolha

#### Coluna F - 💰 SALDO DEVEDOR

- Quanto você ainda deve
- **Começa com:** Valor financiado
- **Diminui:** A cada pagamento
- **Termina com:** R$ 0,00 (quando quitado)
- 📊 **Visual:** Barra azul mostra proporção (quanto menor, mais perto da quitação)

#### Coluna G - 💳 PARCELA (de L4/L5)

- Valor da parcela base
- No PRICE: sempre igual
- No SAC: diminui com o tempo
- **Incluindo:** Apenas amortização e juros (sem seguro ainda)

#### Coluna H - 🔨 AMORTIZAÇÃO

- Quanto de seu pagamento vai realmente reduzir a dívida
- **No PRICE:** Aumenta ao longo do tempo (juros diminuem)
- **No SAC:** É sempre o mesmo (divida prazo)
- **Cor:** Verde claro (mostra intensidade com cores)

#### Coluna I - 📈 JUROS

- Juros cobrados naquele mês
- **No PRICE:** Começam altos, diminuem
- **No SAC:** Começam altos, diminuem rápido
- **Cor:** Vermelho claro
- 💡 **Dica:** Observe como diminuem ao pagar antecipado

#### Coluna J - ➕ AMORT. ADICIONAL ⚡ **(VOCÊ PREENCHE)**

- Quanto de extra você quer pagar naquele mês
- **Deixe em branco:** Usa valor padrão de H6
- **Digite um valor:** Usa esse valor naquele mês
- **Digite 0:** Não paga nada de extra naquele mês
- 💡 **Dica:** Use aqui quando recebe bônus, 13º, ou tem valores sobressalentes

#### Coluna K - 🛡️ SEGURO

- Seguro mensal
- Preenchida automaticamente com base em H5

#### Coluna L - 📅 DATA PGTO ⚡ **(VOCÊ PREENCHE)**

- Data que você pagou a parcela
- **Deixe em branco:** Indica parcela ainda não paga
- **Digite uma data:** Marca quando pagou
- **Formato:** `DD/MM/YYYY` (ex: `15/01/2025`)
- 💡 **Dica:** Coloque a data do comprovante ou extrato

#### Coluna M - ✏️ VALOR PAGO ⚡ **(VOCÊ PREENCHE)**

- Quanto você pagou realmente
- **Deixe em branco:** Indica não pago
- **Digite o valor:** Quanto você pagou
- 💡 **Dica:** Pode ser diferente da parcela se pagou com atraso, desconto, etc.

#### Coluna N - 📊 TOTAL ACUMULADO

- Soma acumulada de tudo que você já pagou
- Preenchida automaticamente
- **Cresce:** Conforme você paga as parcelas

#### Coluna O - 🚦 SITUAÇÃO

- Status da parcela
- **Opções:**
  - ✅ **PAGO:** Você pagou o valor total ou mais
  - ⚠️ **PARCIAL:** Você pagou, mas menos que o devido
  - ⏳ **PENDENTE:** Ainda não pagou
- **Cores:**
  - Verde = Pago ✅
  - Vermelho = Parcial ⚠️
  - Amarelo = Pendente ⏳

---

## 📈 Interpretando os Resultados

### Zona de Totalizadores (Linha 10-11)

#### ✅ PARCELAS PAGAS

- Quantas parcelas você já quitou
- **Exemplo:** 12 (significa 12 parcelas pagas)

#### 💰 TOTAL AMORTIZADO

- Quanto de principal você já pagou
- **Cresce:** Conforme os meses passam
- **Meta:** Chegar ao valor financiado

#### 💸 TOTAL JUROS PAGOS

- Quanto você já pagou de juros
- **SAC vs PRICE:** Será menor no SAC se pagar antecipado

#### 🛡️ TOTAL SEGUROS

- Total de seguro pago até agora

#### 📊 SALDO DEVEDOR ATUAL

- Quanto você ainda deve
- **Começa:** No valor financiado
- **Termina:** Em R$ 0,00
- ⚠️ **Importante:** Pode mudar se pagar antecipado

#### 📈 % QUITADO

- Percentual do financiamento que você já quitou
- **Começa:** 0%
- **Termina:** 100% (quando acabar de pagar)
- 💡 **Dica:** Use para acompanhar progresso visualmente

### Linha de Totais Finais

Na última linha, você verá o total geral de:

- Parcelas pagas
- Amortização total
- Juros totais
- Seguros totais
- Valores pagos

---

## 💡 Dicas Avançadas

### Como Simular Pagamentos Antecipados

1. **Exemplo:** Você quer quitar a parcela 5 em uma data anterior

**Passo a passo:**

```
Linha 18 (mês 5):
- Coluna L: Digite a data que quer pagar (ex: 15/01/2025)
- Coluna M: Digite o valor que quer pagar (ex: 5.000,00)
- Coluna J: Se quiser pagar extra, coloque aqui
```

2. **Resultado:** A situação mudará automaticamente
   - Se pagou >= ao valor da parcela: ✅ PAGO
   - Se pagou < ao valor da parcela: ⚠️ PARCIAL
   - Se não pagou: ⏳ PENDENTE

### Como Usar Amortização Adicional Eficientemente

**Sistema SAC + Amortização adicional = Melhor combinação**

1. No sistema SAC, você já paga mais no início
2. Adicione amortização extra nos primeiros meses
3. Resultado: Juros caem rapidamente, saldo diminui rápido

### Como Acompanhar Mês a Mês

1. **Todo mês:**
   - Na linha do mês, preencha Coluna L (data) e M (valor pago)
   - Coloque J (amortização extra) se tiver disponível

2. **Observe:**
   - Coluna O mostra se está pago/parcial/pendente
   - Coluna N mostra total acumulado pago
   - Totalizadores atuam em tempo real

### Dicas para Economizar em Juros

1. **Aumente amortização nos primeiros anos**
   - Juros são maiores no início
   - Pagar extra cedo = menos juros no total

2. **Compare PRICE vs SAC:**
   - SAC economiza mais juros
   - Mas primeira parcela é maior

3. **Recebeu bônus ou 13º?**
   - Coloque em amortização adicional
   - Pode reduzir anos de financiamento

4. **Use a barra de saldo devedor:**
   - Observe como cai quando você paga extra
   - Motivação visual de progresso

---

## ❓ Perguntas Frequentes

### P: Qual sistema escolher, PRICE ou SAC?

**R:** Depende do seu objetivo:

- **PRICE:** Se prefere parcela fixa e previsível
- **SAC:** Se quer pagar menos juros no total

### P: Posso mudar o sistema depois de preencher dados?

**R:** Sim, mas os cálculos de juros mudarão. Todos os dados da tabela se recalcularão automaticamente.

### P: O que fazer se cometi um erro ao preencher?

**R:** Simples:

1. Clique na célula errada (Coluna J, L ou M)
2. Limpe o valor (Delete)
3. Digite o valor correto
4. Tudo se recalcula automaticamente

### P: Posso adicionar mais linhas para mais meses?

**R:** A planilha já tem 360 meses. Se precisar de mais:

1. Copie uma linha existente
2. Cole abaixo do total
3. Ajuste as fórmulas manualmente

### P: Como saber se estou pagando corretamente?

**R:** Olhe a Coluna O:

- ✅ PAGO = Tudo certo
- ⚠️ PARCIAL = Você pagou menos, há saldo devedor
- ⏳ PENDENTE = Ainda não pagou

### P: Posso usar a planilha para simular diferentes cenários?

**R:** Sim! Faça assim:

1. Salve a planilha com nome diferente (ex: `Cenario_1.xlsx`)
2. Mude os valores de configuração
3. Todos os cálculos atualizam automaticamente
4. Compare os cenários lado a lado

### P: Os valores incluem inflação?

**R:** Não. Esta planilha considera taxa de juros fixa, sem inflação. Para análises futuras, consulte seu banco.

### P: Posso exportar ou imprimir?

**R:** Sim! Qualquer software de planilhas permite:

- **Imprimir:** Ctrl+P (ou Cmd+P)
- **Exportar:** File → Export
- **Compartilhar:** Salve na nuvem (Google Drive, OneDrive, etc.)

### P: A planilha está travando ou muito lenta?

**R:** Se houver muitos dados:

1. Salve em formato `.xlsx` (não `.xls`)
2. Feche e abra novamente
3. Se muito lento, importe para Google Sheets (costuma ser mais rápido)

### P: Como interpretar o "Total de Juros"?

**R:** É a diferença entre:

- **Total a Pagar** (tudo que você vai pagar)
- **Valor Financiado** (o que pegou emprestado)

**Exemplo:**

- Financiou: R$ 500.000,00
- Total a Pagar: R$ 900.000,00
- Total de Juros: R$ 400.000,00

---

## 📞 Suporte e Melhorias

Se encontrar erros ou tiver sugestões:

1. Verifique os valores de entrada (às vezes um número errado causa cascata de erros)
2. Limpe as células com erro e refaça
3. Se o problema persistir, regere a planilha executando `create_excel_imovel.py` novamente

---

## 🎨 Legenda de Cores

| Cor                     | Significado                                  |
| ----------------------- | -------------------------------------------- |
| 🟨 **Amarelo/Laranja**  | Campo para você preencher (entrada de dados) |
| 🟢 **Verde claro**      | Amortização / valores calculados positivos   |
| 🔴 **Vermelho claro**   | Juros (diminui com SAC ao longo do tempo)    |
| 🔵 **Azul claro**       | Saldo devedor (barra visual)                 |
| ✅ **Verde escuro**     | Parcela PAGA                                 |
| ⏳ **Amarelo intenso**  | Parcela PENDENTE                             |
| ⚠️ **Vermelho intenso** | Parcela PARCIAL                              |

---

## ✨ Conclusão

A Tabela de Amortização é uma ferramenta poderosa para:

- 📊 Acompanhar seu financiamento
- 💡 Simular cenários
- 💰 Identificar oportunidades de economia
- 📈 Visualizar progresso

**Dica final:** Atualize a planilha regularmente com seus pagamentos reais. Isso ajuda a manter um controle preciso e identificar se há alguma discrepância com seu banco.

**Boa sorte com seu financiamento! 🏠**

---

_Manual criado para facilitar o uso da Tabela de Amortização de Financiamento Imobiliário_
_Última atualização: Abril 2026_

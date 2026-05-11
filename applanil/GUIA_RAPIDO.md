# ⚡ Guia Rápido - 5 Minutos para Começar

## 📌 Resumo em 5 Passos

### 1️⃣ Gere a Planilha

```bash
python create_excel_imovel.py
```

Será criado: `Imovel_Financiado_PERSONALIZADO.xlsx`

### 2️⃣ Abra no Excel/Google Sheets

Clique duplo no arquivo ou abra com seu programa de planilhas

### 3️⃣ Preencha as Configurações (Seção Roxa)

Você precisa preencher **4 campos obrigatórios:**

| Campo                    | Exemplo  | Notas            |
| ------------------------ | -------- | ---------------- |
| **💰 Valor Financiado**  | 500000   | Em reais, sem R$ |
| **📅 Mês/Ano de Início** | Jan/2025 | Quando começa    |
| **⏳ Prazo (meses)**     | 360      | 360 = 30 anos    |
| **📈 Taxa (% a.m.)**     | 0.75     | Como percentual  |

**Opcionais:**

- **🏦 Sistema:** SAC ou PRICE (padrão: SAC)
- **🛡️ Seguro Mensal:** Deixe em branco = R$ 0
- **📊 Amort. Adicional:** Deixe em branco = sem extra

### 4️⃣ Veja os Cálculos Aparecerem

✨ Tudo se calcula automaticamente!

**Você verá:**

- Parcela mensal (em L4)
- Total a pagar (em L6)
- Total de juros (em L7)

### 5️⃣ Preencha os Pagamentos (Conforme Pagar)

Para cada mês que pagar:

| Coluna | O que Preencher              | Formato                    |
| ------ | ---------------------------- | -------------------------- |
| **J**  | Amortização extra (opcional) | Número ou deixar em branco |
| **L**  | Data que pagou               | DD/MM/YYYY                 |
| **M**  | Quanto pagou                 | Número em reais            |

A situação (✅ PAGO / ⏳ PENDENTE / ⚠️ PARCIAL) **atualiza sozinha** ⚡

---

## 🎯 Exemplo Prático

### Cenário: Você pega R$ 500 mil por 30 anos a 0,75% a.m. (SAC)

**Preenchimento inicial:**

```
D4 (Valor):        500000
D5 (Mês/Ano):      Jan/2025
D6 (Prazo):        360
D7 (Taxa):         0.75%
H4 (Sistema):      SAC
H5 (Seguro):       250 (exemplo)
H6 (Amort Extra):  2000 (exemplo)
```

**Resultado automático:**

```
L4 (Parcela SAC):     R$ 4.231,43 (primeira, depois diminui)
L6 (Total a Pagar):   R$ 1.285.392,28
L7 (Total Juros):     R$ 785.392,28
```

**Depois, você preenche mensalmente:**

```
Mês 1 (Janeiro/2025):
  - J14: deixa em branco (usa padrão de 2000)
  - L14: 15/01/2025 (data que pagou)
  - M14: 6481.43 (parcela + seguro)
  → O14 automaticamente marca: ✅ PAGO
```

---

## 🎨 Cores = O Que Significa

```
🟨 AMARELO/LARANJA    → VOCÊ PREENCHE AQUI
🟢 VERDE CLARO        → Dados calculados (não mexer)
🔵 AZUL CLARO         → Saldo devedor
✅ VERDE ESCURO       → Parcela paga
⏳ AMARELO FORTE      → Parcela pendente
⚠️ VERMELHO FORTE     → Parcela paga parcialmente
```

---

## ❓ Dúvidas Rápidas

**P: Qual sistema escolher?**

- SAC = Paga menos juros total (recomendado) 👍
- PRICE = Parcela fixa (mais previsível)

**P: Posso mudar depois?**

- Sim! Tudo se recalcula automaticamente

**P: Esqueci de preencher um mês?**

- Volte e preencha. A planilha atualiza tudo

**P: Recebi bônus e quero pagar extra?**

- Coloque em J (coluna de amortização adicional)
- Saldo devedor diminui mais rápido 💪

**P: Quero comparar cenários?**

- Salve com nome diferente
- Mude valores de configuração
- Compare os totais

---

## 📊 Dados em Tempo Real

A planilha **atualiza tudo sozinha**:

- Linha 10-11: Totalizadores (parcelas pagas, saldo, juros, etc.)
- Colunas N: Total acumulado pago
- Coluna O: Situação de cada parcela

Basta você preencher as colunas amarelas! ⚡

---

## 🚀 Próximos Passos

1. **Comece agora:** Gere a planilha e preencha as 4 informações iniciais
2. **Leia o manual completo:** Veja `MANUAL_USUARIO.md` para detalhes
3. **Acompanhe mensalmente:** Adicione seus pagamentos conforme pagar
4. **Simule:** Teste amortizações extras para ver o impacto

---

**Dica:** Atualize a planilha todo mês com seus pagamentos reais. Você verá o progresso em tempo real! 📈

Qualquer dúvida, consulte o **MANUAL_USUARIO.md** completo! 📘

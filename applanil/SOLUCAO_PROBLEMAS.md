# 🔧 Dicas de Solução de Problemas

## ⚠️ Problemas Comuns e Soluções

### Problema 1: "A planilha mostra #DIV/0! ou #N/A"

**Causa:** Um dos campos de configuração está vazio ou com formato incorreto

**Solução:**

1. Verifique a **seção roxa (linhas 4-7)**
2. Preencha todos os 4 campos obrigatórios:
   - D4: Valor Financiado (número sem R$)
   - D5: Mês/Ano (ex: Jan/2025)
   - D6: Prazo (número de meses)
   - D7: Taxa de juros (percentual)
3. Se ainda mostrar erro, delete e digite novamente
4. Execute novamente: `python create_excel_imovel.py`

---

### Problema 2: "As fórmulas não estão calculando"

**Causa:** Planilha pode estar em modo de visualização ou cálculo manual

**Solução (Excel):**

1. Vá em **Fórmulas** → **Opções de Cálculo** → **Automático**
2. Ou pressione **Ctrl + Shift + F9** (recalcular tudo)

**Solução (Google Sheets):**

1. Vá em **Arquivo** → **Configurações** → **Cálculo**
2. Marque **Recálculo automático**

**Solução (LibreOffice):**

1. Vá em **Ferramentas** → **Opções** → **LibreOffice Calc** → **Calcular**
2. Marque **Recalcular automaticamente**

---

### Problema 3: "Não consigo preencher as células amarelas"

**Causa:** A planilha pode estar protegida

**Solução (Excel):**

1. Vá em **Ferramentas/Revisão** → **Desproteger Planilha**
2. Se pedir senha, deixe em branco e confirme

**Solução (Google Sheets):**

1. Vá em **Dados** → **Proteger intervalos**
2. Desmarque proteção se houver

**Solução (LibreOffice):**

1. Vá em **Ferramentas** → **Proteger Documento** → **Desproteger Planilha**

---

### Problema 4: "A data de vencimento não aparece corretamente"

**Causa:** Formato de data diferente do esperado

**Solução:**

1. Clique em uma célula da coluna D (vencimento)
2. Verifique o formato: deve ser **DD/MM/YYYY**
3. Se não estiver, formate novamente:
   - **Excel:** Clique direito → Formatar Células → Data → DD/MM/YYYY
   - **Google Sheets:** Formato → Número → Data (escolha DD/MM/YYYY)

---

### Problema 5: "Os valores aparecem como #### ou muito pequenos"

**Causa:** Coluna muito estreita para o valor

**Solução:**

1. Posicione o cursor na divisória entre as letras das colunas
2. Clique e arraste para aumentar a largura
3. Ou clique duplo na divisória para auto-ajustar

---

### Problema 6: "Preenchi um valor e a planilha inteira desapareceu"

**Causa:** Provavelmente um pem dos atalhos acidentalmente deletou conteúdo

**Solução:**

1. Pressione **Ctrl + Z** (desfazer)
2. Repita até voltar ao estado anterior
3. Se não funcionar, execute novamente: `python create_excel_imovel.py`

---

### Problema 7: "Os totalizadores (linha 10-11) não atualizam"

**Causa:** Planilha não recalculou automaticamente

**Solução:**

1. Clique em uma célula dos totalizadores
2. Pressione **F9** (recalcular)
3. Se não funcionar, salve e feche/abra novamente

---

### Problema 8: "Não consigo abrir o arquivo"

**Causa:** Programa instalado não suporta Excel ou arquivo corrompido

**Solução:**

1. Tente abrir com outro programa:
   - **Excel** (recomendado)
   - **Google Sheets** (versão online, gratuita)
   - **LibreOffice Calc** (gratuito)
2. Se mesmo assim não abrir, regere o arquivo:
   ```bash
   python create_excel_imovel.py
   ```

---

### Problema 9: "A cor de fundo não aparece ou está incorreta"

**Causa:** Configurações de exibição de cores

**Solução (Excel):**

1. **Arquivo** → **Opções** → **Avançado**
2. Procure por "Imprimir" → **Imprimir cores e imagens**
3. Marque essa opção

**Solução (Google Sheets):**

1. Cores devem aparecer automaticamente
2. Se não aparecer, tente em outro navegador

---

### Problema 10: "Recebi um erro ao executar o Python"

**Mensagens comuns:**

#### "ModuleNotFoundError: No module named 'openpyxl'"

```
Solução: Instale a dependência
pip install openpyxl
```

#### "FileNotFoundError: create_excel_imovel.py"

```
Solução: Certifique-se que o arquivo existe no diretório
cd c:\Users\kauan\Downloads\applanil
python create_excel_imovel.py
```

#### "Permission denied" ou "Access denied"

```
Solução: O arquivo pode estar aberto em outro programa
1. Feche a planilha anterior
2. Execute novamente
```

---

## 🔍 Verificações Rápidas

### Checklist antes de reclamar que algo está errado:

- [ ] Os 4 campos obrigatórios estão preenchidos?
- [ ] O formato de data está correto (DD/MM/YYYY)?
- [ ] A taxa de juros é um percentual (ex: 0,75%, não 75%)?
- [ ] O prazo é um número inteiro (360, não 30 anos)?
- [ ] A planilha está em modo de cálculo automático?
- [ ] Nenhuma coluna está muito estreita (#### aparecendo)?
- [ ] Pressionei F9 ou ctrl+shift+F9 para recalcular?
- [ ] A planilha não está protegida?

Se todas estão OK e ainda há problema, siga as soluções acima.

---

## 💾 Dicas de Backup e Segurança

### Faça Backups Regulares

```
1. Todo mês, faça uma cópia:
   Copie "Imovel_Financiado_PERSONALIZADO.xlsx"
   Cole e renomeie: "Imovel_Financiado_Janv2025_BACKUP.xlsx"

2. Guarde em um local seguro:
   - Google Drive
   - OneDrive
   - Pen drive
```

### Proteja contra Alterações Acidentais

```
Excel:
1. Ferramentas → Proteger Planilha
2. Digite uma senha simples
3. Só deixe desbloqueadas as colunas J, L, M

Google Sheets:
1. Dados → Proteger intervalos
2. Selecione apenas colunas J, L, M
3. Configure permissões
```

---

## 📞 Quando Pedir Ajuda

Se nada funcionou:

1. **Descreva o erro:** O que vê na tela?
2. **Qual sistema:** Excel, Google Sheets, LibreOffice?
3. **Qual valor causou:** A qual campo estava preenchendo?
4. **Qual versão:** Qual versão do programa está usando?
5. **Screenshots:** Tire uma captura de tela mostrando o erro

---

## 🚀 Otimizações para Performance

### Se a planilha está lenta:

1. **Salve em .xlsx** (não .xls)
2. **Feche outras abas/programas**
3. **Use Google Sheets** (geralmente mais rápido)
4. **Limite o scroll** (não passe a linha 420)
5. **Desative filtros** se tiver ativado

### Se está muito lenta mesmo assim:

1. Crie uma cópia apenas com os dados que precisa
2. Ou use em Google Sheets que é otimizado para nuvem

---

## 🔄 Regenerar a Planilha do Zero

Se nada funcionar, comece de novo:

```bash
# 1. Abra Terminal/PowerShell
# 2. Vá para a pasta
cd c:\Users\kauan\Downloads\applanil

# 3. Ative o ambiente virtual (se houver)
.\.venv\Scripts\Activate

# 4. Reinstale a dependência
pip install openpyxl --upgrade

# 5. Execute o script
python create_excel_imovel.py

# 6. Abra o novo arquivo gerado
Imovel_Financiado_PERSONALIZADO.xlsx
```

---

## 📧 Melhorias Futuras Planejadas

- [ ] Suporte a taxas variáveis
- [ ] Gráficos de amortização
- [ ] Relatório PDF automático
- [ ] Integração com calculadora de taxas
- [ ] Versão com inflação
- [ ] Simulador de cenários lado a lado

---

Esperamos que este guia tenha resolvido sua dúvida! 😊

Se precisar de mais ajuda, consulte o **MANUAL_USUARIO.md** completo.

**Boa sorte com seu financiamento!** 🏠💪

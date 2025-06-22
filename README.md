
# 📊 Simulador de Declaração de Imposto de Renda - Excel

Este é um projeto pessoal desenvolvido no Excel que simula o processo de entrega da Declaração de Imposto de Renda para Pessoa Física (IRPF).

O objetivo é facilitar o planejamento tributário, permitindo ao usuário visualizar de forma clara sua situação fiscal, de maneira totalmente dinâmica, interativa e com interface amigável.

---

## 📂 Estrutura do Projeto

O aplicativo conta com **4 abas principais** e **1 aba de apoio**, detalhadas abaixo:

### 🧑 Titular
- Cadastro de **dados pessoais** do contribuinte (nome, CPF, estado civil, se possui cônjuge, etc.).
- Campo para seleção de **dependente (Sim/Não)**, utilizado no cálculo das deduções.
- Integração com as demais abas para cálculo dinâmico das deduções.

---

### 🏦 Informes
- Lançamento de **rendimentos recebidos de até 3 contas bancárias**, com os seguintes campos:
  - Banco
  - Valor recebido
- Campo para preenchimento manual do **IR Retido na Fonte** (valor informado pelas fontes pagadoras).
- **Gráfico de Pizza (Pie Chart)** mostrando a proporção de rendimento por banco.
- Todos os campos preenchidos pelo próprio declarante.

---

### 📝 Notas
- Registro detalhado de **rendas adicionais**, com os seguintes campos:
  - **Data**
  - **Categoria** (CNPJ / Holerite / Freelance) com lista suspensa.
  - **Valor Recebido**
- Permite acompanhamento de diferentes tipos de recebimentos fora das contas bancárias.

---

### 🧾 Resumo IR
- Exibe um **resumo final** com os principais cálculos:
  - **Rendimento Total**
  - **Total de Deduções**
  - **Base de Cálculo**
  - **Imposto Simulado**
  - **IR Retido**
  - **Diferença**
  - **Status** (Pagar / Restituição / Isento)

#### 🎨 Recursos Especiais no Resumo IR:
- Cálculos dinâmicos utilizando as funções **ÍNDICE + CORRESP + SEERRO + SE aninhados**.
- Cores de destaque diferentes para cada tipo de **Status**.
- Fórmulas automatizadas para cada faixa de renda segundo a tabela oficial da Receita.

---

### 📌 Aba Auxiliar (Tabela de Apoio)
- Contém a **tabela com as faixas de Imposto de Renda**, incluindo:
  - Limite Inferior
  - Limite Superior
  - Alíquota
  - Parcela a Deduzir
- Estruturada para ser usada nos cálculos com **busca dinâmica de faixa**.

---

## 🛠️ Recursos Extras Utilizados

- **Menus laterais interativos:**  
  Todas as abas possuem um **menu lateral esquerdo** com botões para **navegação rápida** entre as seções.

- **Botões de Navegação:**  
  Incluídos botões **"Próximo"** e **"Avançar"** entre as abas.

- **VBA (Visual Basic for Applications):**  
  Pequeno código VBA utilizado para **alinhar e centralizar o ícone do LinkedIn** nas abas.

- **Formatação de células:**  
  Todas as células foram formatadas com atenção ao layout e à usabilidade:  
  Moeda, datas, listas suspensas, e mensagens de orientação ao usuário.

- **Prints das telas:**  
  Foram tirados **screenshots (prints)** de todas as abas e etapas para futura documentação visual.

---

## ✅ Objetivo Final
Este projeto visa:

- Treinar recursos avançados de Excel como **fórmulas dinâmicas**, **gráficos**, **VBA básico**, **listas suspensas** e **formatação condicional**.
- Oferecer ao usuário uma **experiência completa de simulação** de entrega de IRPF de forma simples e eficiente.

---

## 📸 Prints
segue em anexo a pasta imagens com mais prints
---

## 📅 Status do Projeto
✔️ Concluído com sucesso  
✔️ Em uso para fins de estudo e planejamento pessoal  
✔️ Aberto para futuras melhorias

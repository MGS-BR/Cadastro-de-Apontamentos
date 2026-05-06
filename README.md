# 📊 Automação de Lançamento de Apontamentos

![Python](https://img.shields.io/badge/Python-3.14-blue)
![Status](https://img.shields.io/badge/status-em%20desenvolvimento-yellow)
[![License](https://img.shields.io/badge/license-MIT-green)](https://github.com/MGS-BR/Cadastro-de-Apontamentos/blob/main/LICENSE)
[![Python CI](https://github.com/MGS-BR/Cadastro-de-Apontamentos/actions/workflows/ci.yml/badge.svg)](https://github.com/MGS-BR/Cadastro-de-Apontamentos/actions/workflows/ci.yml)
[![Build EXE](https://github.com/MGS-BR/Cadastro-de-Apontamentos/actions/workflows/build.yml/badge.svg)](https://github.com/MGS-BR/Cadastro-de-Apontamentos/actions/workflows/build.yml)

Aplicação desktop desenvolvida em Python para automatizar o lançamento de apontamentos de funcionários no sistema **Contmatic Phoenix**, utilizando leitura de planilhas Excel e automação de interface gráfica.

---

## 🚀 Objetivo

Eliminar o trabalho manual e repetitivo de digitação de eventos e valores no sistema Contmatic Phoenix, aumentando a produtividade e reduzindo erros operacionais.

---

## 🛠️ Tecnologias Utilizadas

* Python 3
* Tkinter (interface gráfica)
* Pandas (leitura e manipulação de planilhas)
* PyAutoGUI (automação de mouse e teclado)
* tkcalendar (seleção de datas)
* keyboard (hotkeys)

---

## 📦 Funcionalidades

* 📁 Seleção de planilha Excel
* 📅 Definição de período (data inicial e final)
* ▶️ Execução automatizada dos lançamentos
* 📊 Barra de progresso
* 🖥️ Console integrado para logs em tempo real
* ⛔ Interrupção de execução com tecla `ESC`
* ⚠️ Confirmação entre funcionários
* 🔄 Tratamento de erros por linha

---

## 📄 Estrutura da Planilha

A planilha deve conter, a partir da linha 4 (`header=3`):

| Código | Nome | Evento | Ref./Valor |
| ------ | ---- | ------ | ---------- |
| 123    | João | 101    | 150.00     |

### Regras:

* Linhas sem **Nome** ou **Ref./Valor** são ignoradas
* Os dados são ordenados automaticamente por **Código**

---

## ▶️ Como Usar

1. Execute o script:

```bash
python app.py
```

2. Na interface:

* Selecione a planilha
* Escolha a data inicial e final
* Clique em **Executar**

3. Após iniciar:

* Abra o sistema Contmatic Phoenix
* Aguarde o início da automação (6 segundos)

---

## ⚠️ Importante

### 🔧 Configuração de Tela

O sistema utiliza coordenadas fixas de clique (`pyautogui`), portanto:

* A resolução de tela deve ser compatível com a utilizada no desenvolvimento
* O sistema Contmatic deve estar aberto e visível
* Não mova o mouse durante a execução

---

## ⛔ Interrupção de Emergência

Você pode parar a execução de duas formas:

* Pressionando **ESC**
* Movendo o mouse para o canto superior esquerdo (FAILSAFE do PyAutoGUI)

---

## 🧠 Lógica da Automação

Para cada linha da planilha:

1. Seleciona o funcionário (se necessário)
2. Cria um novo lançamento
3. Preenche:

   * Evento
   * Valor
   * Datas
4. Grava o lançamento
5. Solicita confirmação ao trocar de funcionário

---

## 🐞 Tratamento de Erros

* Erros por linha são exibidos no console e ignorados
* Execução continua normalmente
* Interrupções mostram alertas ao usuário

---

## 🚧 Roadmap (A implementar)

Funcionalidades e melhorias planejadas para evolução do projeto:

### 🔧 Melhorias Técnicas

* [ ] Mapeamento dinâmico de coordenadas (reduzir dependência de resolução fixa)
* [ ] Sistema de calibração automática de cliques
* [ ] Melhor tratamento de exceções com detalhamento por tipo de erro

### ⚙️ Interface e Usabilidade

* [ ] Tela de configurações (⚙ já existente, mas não implementada)
* [ ] Feedback visual mais claro de progresso (% e status atual)
* [ ] Possibilidade de pausar e continuar execução

---

## 📷 Observações

Essa automação depende diretamente da interface gráfica do sistema Contmatic Phoenix. Qualquer alteração visual no sistema pode impactar o funcionamento.

---

## 👨‍💻 Autor

Desenvolvido para automação de rotinas administrativas e ganho de eficiência operacional.

---

## 📄 Licença

Este projeto está sob a licença MIT.

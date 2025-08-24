# 🧾 Extrator de Dados de Boletos v1
Uma aplicação de desktop desenvolvida em Python para extrair informações importantes de boletos e faturas em PDF, como linha digitável, valor, vencimento e QR Code PIX.

## ✨ Funcionalidades

- Interface Gráfica amigável construída com Tkinter.
- Extração de dados de boletos bancários e contas de consumo.
- Leitura de QR Code com filtro inteligente para PIX.
- Lógica de fallback para extrair dados do QR Code se não forem encontrados no texto.
- Exportação dos resultados para planilhas Excel (.xlsx).



###  EXE - Versão Pronta para Uso (Windows) 🚀

Para facilitar a vida de usuários que não são desenvolvedores, eu preparei uma versão executável (`.exe`) do programa. Ela está pronta para rodar no Windows com um simples duplo clique, sem a necessidade de instalar Python ou qualquer outra dependência.

Se esta ferramenta foi útil para você, considere apoiar o desenvolvimento contínuo do projeto com uma pequena contribuição. Isso me ajuda a dedicar mais tempo para criar novas funcionalidades e melhorias!

O processo é simples e automatizado: sua contribuição libera o download imediato do arquivo `.zip` contendo o programa.

➡️ **[Clique aqui para obter a versão .exe via Ko-fi (Pague o que quiser!)](https://ko-fi.com/s/98332f2457)

Sua contribuição, de qualquer valor, é o que mantém projetos como este vivos e evoluindo. Muito obrigado pelo apoio!


## 🚀 Como Usar (Para Desenvolvedores)

1. Clone o repositório.
2. Crie um ambiente virtual: `python -m venv .venv`
3. Ative o ambiente e instale as dependências: `pip install -r requirements.txt`
4. Execute o programa: `python organizador_boletos_v1.py`




## 🛠️ Tecnologias Utilizadas

- Python
- Tkinter
- Pandas
- PyMuPDF (fitz)
- Pyzbar
- Pillow

---

### 🤖 Sobre o Desenvolvimento e o Uso de IA

Este projeto foi idealizado e desenvolvido por Albert7z. Durante todo o processo, utilizei ferramentas de Inteligência Artificial, como o Google Gemini, como uma assistente de programação e parceira de "brainstorming".

A IA foi fundamental para:
* Acelerar a escrita de código repetitivo (boilerplate).
* Depurar (debug) erros complexos e exceções.
* Sugerir e refinar algoritmos e expressões regulares (Regex).
* Explorar diferentes abordagens para os desafios que surgiram no caminho.

Toda a arquitetura, a lógica principal e as decisões finais do projeto foram conduzidas por mim, com a IA servindo como uma poderosa ferramenta para aumentar a produtividade e a qualidade do código final.

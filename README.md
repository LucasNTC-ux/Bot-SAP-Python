## Olá, este script foi desenvolvido em Python para comunicação a uma transação SAP, feita para rodar e realizar um relatório de demanda.

## Para que o script funcione perfeitamente, precisará instalar algumas bibliotecas (libs) no seu computador, são essas:
1. subprocess
2. sys
3. keyboard
4. pyautogui
5. time
6. tkinter

Para instalar, é bem fácil, basta abrir o terminal e digitar:
```
pip install subprocess
```
Você pode encontar mais lib's em:
[PyPI](https://pypi.org/)


## Para que ele vire um executável, siga os passos:
1. PRIMEIRO: vamos baixar a biblioteca pyinstaller.
Escreva em seu terminal:
```
pip install pyinstaller
```
2. SEGUNDO: após isso, no terminal aonde está seu código.
Digite o seguinte comando:
```
pyinstaller -onefile NOME_DO_SEU_ARQUIVO.py
```
## PRONTO, agora ele irá criar uma pasta "dist" dentro da pasta, nela você irá encontrar o .exe da sua aplicação!

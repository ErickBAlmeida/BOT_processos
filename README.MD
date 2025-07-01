# Automação de Procesos Eletrônicos - Script de Testes

Este projeto automatiza o processo de anexar petições e procurações no sistema processos eletrônicos utilizando Python, Selenium e Pywinauto.

## Funcionalidades

- Acessa o sistema de processos judiciais do estado configurado.
- Navega automaticamente pelo menu do sistema.
- Busca processos a partir dos arquivos de petição em uma pasta específica.
- Anexa petições e procurações em processos selecionados.
- Detecta processos sigilosos e casos onde o advogado já faça parte do processo.
- Move petições processadas para uma pasta separada na área de trabalho.
- Exibe notificações de status e erros via Toast.

## Pré-requisitos

- Python 3.11+
- Google Chrome instalado
- [Selenium](https://pypi.org/project/selenium/)
- [Pywinauto](https://pypi.org/project/pywinauto/)
- [win10toast](https://pypi.org/project/win10toast/)
- ChromeDriver compatível com sua versão do Chrome

## Instalação

```sh
pip install selenium pywinauto win10toast
<h1 align="center">
📄<br>README - Projeto Automação Web Compras Online
</h1>

## Índice 

* [Descrição do Projeto](#descrição-do-projeto)
* [Funcionalidades e Demonstração da Aplicação](#funcionalidades-e-demonstração-da-aplicação)
* [Pré requisitos](#pré-requisitos)
* [Execução](#execução)
* [Bibliotecas](#bibliotecas)

# Descrição do projeto
> Este repositório é meu projeto Python de automação web (web-scrapping) e envio de e-mail com informações para as compras de produtos online. A partir da base de dados com os produtos de interesse, o objetivo do projeto foi encontrar as melhores ofertas no Google Shopping e, com as informações coletadas, enviar um e-mail com os valores e os links de compras dos produtos no corpo do e-mail e no arquivo anexado.

# Funcionalidades e Demonstração da Aplicação

E-mail enviado com as informações dos melhores produtos encontrados no Google Shopping:<br>
![Screenshot_1](https://user-images.githubusercontent.com/128300382/228635114-e01388d6-a046-46d3-8a09-016512a88636.png)

## Pré requisitos

* Sistema operacional Windows
* IDE de python (ambiente de desenvolvimento integrado de python)
* Base de dados (arquivo Excel)
* Navegador Google Chrome (para o web-scrapping)

## Execução

O código, ao ser executado, realiza uma automação no navegador Google Chrome, buscando no Google Shopping as ofertas de venda dos produtos elencados na base de dados "buscas.xlsx". Após o web-scrapping, as informações são armazenadas e exportadas em uma planilha Excel "Ofertas Google Shopping (revisão).xlsx". A partir dessa base de dados criada, um e-mail enviado com as informações necessárias para a compra dos produtos.

## Bibliotecas

* <strong>pandas:</strong> bibliotecas de integração de arquivos excel, csv e outros, possibilitando análise de dados<br>
* <strong>os:</strong> biblioteca de integração de arquivos e pastas do computador<br>
* <strong>win32com.client:</strong> biblioteca de integração dos aplicativos Windows, no caso, do Outlook<br>
* <strong>selenium, webdriver_manager:</strong> biblioteca de automação web<br>
* <strong>pprint:</strong> biblioteca de visualização de dicionários complexos<br>
* <strong>datetime:</strong> biblioteca que permite a utilização de datas e horários<br>
* <strong>time:</strong> biblioteca que permite o gerenciamento do tempo na execução do código<br>

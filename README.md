# Webscraping2
Automação de Webscraping para Demandas Internas da Empresa 2
Este projeto foi desenvolvido para atender a uma demanda interna da empresa onde trabalho. Esta aplicação difere da anterior, pois, além de ser em um sistema diferente, ela também possui uma interface e realiza o login automaticamente. Ela consulta se o cliente está aprovado ou reprovado.

Funcionalidades:
•	A aplicação abre a interface solicitando dados de login, senha e token de autenticação.
•	Abre o site e faz o login.
•	Em seguida, realiza consultas no sistema com base em números obtidos de um arquivo de dados interna(que é inserido manualmente e fica no mesmo local de arquivo do programa).
•	A aplicação faz diversos clicks e faz um total de 4 mudanças de páginas, preenche formulários e coleta 1 informação que é salva em um arquivo Excel.
•	Esse processo é repetido enquanto tiver números no arquivo de dados interno.

Estrutura do Código:
O código está organizado da seguinte forma:
1.	Importação de bibliotecas e módulos no início do script.
2.	Construção da interface básica, apenas solicitando login, senha e token.
3.	Definição de variáveis para gerar a data atual no formato necessário para a planilha resultante.
4.	Configuração para abrir o site da empresa.
5.	Implementação de uma função para realizar o login.
6.	Leitura de um arquivo com os números a serem consultados e criação de uma lista.
7.	Utilização de laços de repetição para realizar o processo de automação em loop.
obs: Dentro do laço tem um try except que está ali para identificar linhas que não são localizadas no sistema.

Vídeo da aplicação sendo executada:
Youtube: https://youtu.be/Rv83bVlCQoY
Linkedin: 

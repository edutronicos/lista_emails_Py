Basicamente o código faz:

1. Abre o Microsoft Outlook.
2. Seleciona um dos e-mails que você usa.
3. Seleciona a pasta dentro desse e-mail (Caixa de entrada, Lixeira, etc.)
4. Le os e-mail não lidos.
5. Procura dentro do corpo da mensagem por e-mails.
6. Retorna uma lista com todos os e-mails encontrados em um arquivo txt.

O Problema era:

O setor de RH da empresa, dispara vários holerites para os funcionários cadastrados na mesma.
Em alguns cadastros os e-mails estavam incorretos.
Então o servidor "Google" retorna um e-mail informando que não foi possivel entregar a mensagem ao destinatario "pessoa@email.com.br".
O retorno era de 200 a 300 e-mails.
Para atualizar tinha que abrir e-mail por e-mail, anotar o e-mail incorreto e informar ao rh para entrar em contato com o funcionario e atualizar.
Executando o código em 6 segundos retorna a lista com todos os e-mails.

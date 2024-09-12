*Envio Automatizado de Mensagens no WhatsApp:*
  Este projeto é uma automação para enviar mensagens pelo WhatsApp Web usando Python. 
  Ele lê informações de uma planilha do Excel, gera mensagens personalizadas e as envia para os contatos listados. 
  A automação utiliza bibliotecas como openpyxl, pyautogui e webbrowser.

*Funcionalidades:*
 -Leitura de Dados: O script lê uma planilha do Excel (clientes.xlsx) que contém os nomes e números de telefone dos destinatários.
 -Geração de Mensagens: Cria mensagens personalizadas para cada contato.
 -Envio Automatizado: Abre o WhatsApp Web, envia a mensagem e fecha a aba.
 -Tratamento de Erros: Registra falhas no envio e gera um arquivo de log (erros.csv) com os detalhes dos contatos para os quais a mensagem não pôde ser enviada.
 
*Requisitos:*
 -Certifique-se de que as seguintes bibliotecas estão instaladas:
    openpyxl - Para manipulação de planilhas Excel.
    pyautogui - Para automação da interface gráfica.
    webbrowser - Para abrir URLs no navegador.
    
*Você pode instalar as dependências necessárias usando pip:*
 -pip install openpyxl pyautogui

*Configuração:*
  -Planilha de Dados:
    Prepare um arquivo Excel chamado clientes.xlsx com duas colunas: Nome e Telefone.
    Certifique-se de que o nome da aba é Sheet1.

  -Imagens para PyAutoGUI:
    Adicione uma imagem chamada seta1.png que representa o botão de enviar no WhatsApp Web. Essa imagem deve ser uma captura de tela do botão de envio.

  -Ajustes:
    Ajuste o tempo de espera (sleep) conforme necessário, dependendo da velocidade da sua conexão com a internet e do desempenho da sua máquina.

*Notas:*
  Tempo de Espera: O script usa sleep para garantir que a página do WhatsApp Web e os elementos estejam totalmente carregados antes de interagir. Ajuste os valores conforme necessário.
  Tratamento de Erros: O script captura e registra erros durante o envio de mensagens, salvando os detalhes em um arquivo erros.csv.
  
Contribuições:
  Se você deseja contribuir para este projeto, sinta-se à vontade para enviar pull requests ou abrir problemas para melhorias.

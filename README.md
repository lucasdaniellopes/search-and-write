# Pesquisa de Planilhas no Google Drive 📊🔍

Este script em Python realiza uma pesquisa em planilhas do Google Drive, filtrando resultados com base em um valor fornecido pelo usuário. Os resultados da pesquisa são adicionados a uma nova planilha.

## 🚀 Requisitos

Antes de executar o script, certifique-se de ter o Python instalado em seu ambiente. Além disso, instale as dependências necessárias usando o arquivo `requirements.txt`. Você pode instalar as dependências executando o seguinte comando:

pip install -r requirements.txt

⚙️ Configuração

Crie uma conta de serviço no Google Cloud Console.

Acesse o Google Cloud Console.
Crie um novo projeto ou selecione um projeto existente.
No menu de navegação à esquerda, vá para "IAM & Admin" > "Service accounts".
Clique em "Create Service Account" e siga as instruções para criar uma nova conta de serviço.
Atribua a função "Editor" à conta de serviço.
Crie e faça download de uma chave JSON para a conta de serviço e salve-a como chave-da-conta-de-servico.json no diretório do projeto.
Compartilhe a pasta do Google Drive que contém suas planilhas com a conta de serviço.

No Google Drive, clique com o botão direito na pasta que contém suas planilhas.
Selecione "Compartilhar" e adicione o endereço de e-mail da conta de serviço que você criou.
Conceda permissões de leitura/escrita conforme necessário.

🚀 Utilização

Execute o script main.py.

Abra um terminal no diretório do projeto.
Execute o comando python main.py.
Siga as instruções para digitar o valor a ser procurado e o mês desejado.
Aguarde enquanto o script pesquisa e adiciona os resultados à nova planilha.

O progresso será exibido no terminal.
O link da nova planilha será exibido no final da execução.

Você pode abrir a nova planilha diretamente a partir deste link.

🖥️ Interface Gráfica

O script também possui uma interface gráfica simples. Para usar a interface gráfica, execute o script gui.py. A interface inclui uma área de texto que exibe os resultados da pesquisa e um contador de limites de cota excedidos.

📝 Observações
O script respeita os limites de cota do Google Drive API, com pausas entre as solicitações para evitar exceder as cotas.
Em caso de erro 429 (limite de cota excedido), o script faz uma pausa e retenta a operação.

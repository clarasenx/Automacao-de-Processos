# Automação de Processos com Python
Este projeto consiste em um script Python que automatiza a geração de relatórios de vendas diárias e anuais para diferentes lojas, além de enviar esses relatórios por email para os gerentes das lojas e para a diretoria.

### Projeto feito com ajuda da Hashtag treinamentos no curso de Python Impressionador
Projeto com foco em trabalhar conteúdos passados ao longo do curso. Com uma pequena diferença do projeto original, que trabalha a integração do email utilizando o outlook(e a biblioteca win32com.client), nesta versão utilizo o smtplib que envia emails utlizando o protocolo padrão para envio de e-mails através da internet que pode rodar de forma independente e em diferentes plataformas(Linux, Windows, macOS e servidores sem Outlook).

## Bibliotecas utilizadas
- pandas
- pathlib
- smtplib
- os
- email.mime
- python-dotenv

## Funcionalidades
- Processamento de Dados: Importa e processa dados de vendas, lojas e emails a partir de arquivos Excel e CSV.
- Geração de Relatórios: Cria relatórios individuais para cada loja com indicadores de desempenho diários e anuais.
- Envio de Emails: Envia emails automatizados para os gerentes das lojas com os relatórios anexados, bem como um email para a diretoria com o ranking das lojas.
- Backup de Arquivos: Salva os relatórios gerados em pastas de backup organizadas por loja e data.

## Observações
- Segurança:
Não compartilhe a senha do email ou outras informações sensíveis. Utilize variáveis de ambiente para armazenar informações confidenciais(.env).

## Passo a passo para utilização do código
1 - Alterar os caminhos para a base de dados desejada.

![image](https://github.com/user-attachments/assets/3979471c-5355-47f6-964f-bce49ce359cb)
Da pasta onde os arquivos criados irão ficar também!

![image](https://github.com/user-attachments/assets/e3dc1a73-f5a7-4b70-be56-7e7fa9af8afd)

2 - Definir metas de acordo com o seu negócio.

![image](https://github.com/user-attachments/assets/172ccd5d-8aab-45a4-a2a1-9418ab3a147d)

3 -  Dentro da função enviarRelatorioLoja e enviarRelatorioDiretoria é preciso alterar os emails de destinatário e remetente, além de mudar a mensagem de texto do relatório em si e o seu cabeçalho. Dentro dessas funções é preciso mudar os arquivos anexados também.

4 - Ainda dentro dessas funções na parte de conexão com o servidor é preciso criar um arquivo .env e garantir que ele esteja no .gitignore. Dentro do .env você cria a váriavel PASSWORD, assim:

![image](https://github.com/user-attachments/assets/26a19446-2129-48c4-b6fd-1f339738d7e4)

* Para conseguir esse tipo de senha(Senha de app) é preciso ir no gerenciamento de conta do Gmail e ir em Senhas de App e criar uma nova senha. 


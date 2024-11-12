# 1- Importar bibliotecas e as bases de dados
import pandas as pd
import pathlib
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from IPython.display import display
from dotenv import load_dotenv

# mudar variáveis/caminhos para base de dados desejada
emails = pd.read_excel(r'Bases de dados\Emails.xlsx') 
lojas = pd.read_csv(r'Bases de dados\Lojas.csv',  encoding='latin-1', sep=';')
vendas = pd.read_excel(r'Bases de dados\Vendas.xlsx')

# 2- Criar tabela para cada loja e definir o dia do indicador

# incluir nome da loja em vendas
vendas = vendas.merge(lojas, on='ID Loja')

# criar um dicionario com as tabelas de cada loja
dicLojas = {}
for loja in lojas['Loja']:
    dicLojas[loja] = vendas.loc[vendas['Loja']==loja, :]

# criar os indicadores
diaIndicador = vendas['Data'].max()

# 3- Salvar planilhas na pasta de backup
# identificar se a pasta já existe
caminhoBackup = pathlib.Path("Backup Arquivos Lojas")
arquivosPastaBackup = caminhoBackup.iterdir()

listaNomesBackup = [arquivo.name for arquivo in arquivosPastaBackup]

for loja in dicLojas:
    if loja not in listaNomesBackup:
        novaPasta = caminhoBackup/loja
        novaPasta.mkdir()
    
    # salvar dentro da pasta
    nomeArquivo = f'{diaIndicador.month}_{diaIndicador.day}_{loja}.xlsx'
    localArquivo = caminhoBackup/loja/nomeArquivo
    
    dicLojas[loja].to_excel(localArquivo)

# 4- Calcular o indicador para cada loja
# definicao de metas 
metaFaturamentoDia = 1000
metaFaturamentoAno = 165000
metaQtdProdutosDia = 4
metaQtdProdutosAno = 120
metaTicketMedioDia = 500
metaTicketMedioAno = 500

# o for acessa cada loja e envia o relatorio de acordo com seus dados
for loja in dicLojas:
    vendasLoja = dicLojas[loja]
    vendasLojaDia = vendasLoja.loc[vendasLoja['Data']==diaIndicador, :]

    # faturamento
    faturamentoAno = vendasLoja['Valor Final'].sum()
    faturamentoDia = vendasLojaDia['Valor Final'].sum()

    # diversidade de produtos
    qtdProdutosAno = len(vendasLoja['Produto'].unique())
    qtdProdutosDia = len(vendasLojaDia['Produto'].unique())
    
    # ticket medio
    valorVenda = vendasLoja.groupby('Código Venda').sum(numeric_only=True)
    ticketMedioAno = valorVenda['Valor Final'].mean()

    valorVendaDia = vendasLojaDia.groupby('Código Venda').sum(numeric_only=True)
    ticketMedioDia = valorVendaDia['Valor Final'].mean()

    # 5- Enviar email para o gerente

    def enviarRelatorioLoja():
        nome = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
        destinatarios = [nome, "seuEmail@gmail.com"]
        
        mailMessage = MIMEMultipart() #cria o email
        
        mailMessage["From"] = "seuEmail@gmail.com" #remetente
        mailMessage["To"] = ",".join(destinatarios) #destinatario
        
        mailMessage["Subject"] = f'OnePage Dia {diaIndicador.day}/{diaIndicador.month} - Loja {loja}' #cabeçalho
        
        
        if faturamentoDia >= metaFaturamentoDia:
            cor_fat_dia = 'green'
        else:
            cor_fat_dia = 'red'
        if faturamentoAno >= metaFaturamentoAno:
            cor_fat_ano = 'green'
        else:
            cor_fat_ano = 'red'
        if qtdProdutosDia >= metaQtdProdutosDia:
            cor_qtde_dia = 'green'
        else:
            cor_qtde_dia = 'red'
        if qtdProdutosAno >= metaQtdProdutosAno:
            cor_qtde_ano = 'green'
        else:
            cor_qtde_ano = 'red'
        if ticketMedioDia >= metaTicketMedioDia:
            cor_ticket_dia = 'green'
        else:
            cor_ticket_dia = 'red'
        if ticketMedioAno >= metaTicketMedioAno:
            cor_ticket_ano = 'green'
        else:
            cor_ticket_ano = 'red'
        
        # corpo do email
        mailBody = mailMessage.HTMLBody = f'''
        <p>Bom dia, {nome}</p>

        <p>O resultado de ontem <strong>({diaIndicador.day}/{diaIndicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

        <table>
        <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
        </tr>
        <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamentoDia:.2f}</td>
            <td style="text-align: center">R${metaFaturamentoDia:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
        </tr>
        <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtdProdutosDia}</td>
            <td style="text-align: center">{metaQtdProdutosDia}</td>
            <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
        </tr>
        <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticketMedioDia:.2f}</td>
            <td style="text-align: center">R${metaTicketMedioDia:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
        </tr>
        </table>
        <br>
        <table>
        <tr>
            <th>Indicador</th>
            <th>Valor Ano</th>
            <th>Meta Ano</th>
            <th>Cenário Ano</th>
        </tr>
        <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamentoAno:.2f}</td>
            <td style="text-align: center">R${metaFaturamentoAno:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
        </tr>
        <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtdProdutosAno}</td>
            <td style="text-align: center">{metaQtdProdutosAno}</td>
            <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
        </tr>
        <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticketMedioAno:.2f}</td>
            <td style="text-align: center">R${metaTicketMedioAno:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
        </tr>
        </table>

        <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

        <p>Qualquer dúvida estou à disposição.</p>
        <p>Att., Lira</p>
        '''
        
        mailMessage.attach(MIMEText(mailBody, "html"))

        # anexos (pode por quantos quiser)
        attachment = pathlib.Path.cwd()/caminhoBackup/loja/f'{diaIndicador.month}_{diaIndicador.day}_{loja}.xlsx'
        # abre o arquivo, lê e o adiciona ao email
        with open(attachment, "rb") as arquivo:
            mailMessage.attach(MIMEApplication(arquivo.read(), name="Relatorio.xlsx"))
        
        
        # conexão com o servidor
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls() # formato de criptografia para enviar email em segurança
        # conexao o .env para acessar password de forma segura
        load_dotenv()
        password = os.getenv("PASSWORD")
        
        servidor.login(mailMessage["From"], password) # conecta o servidor ao seu email
        servidor.send_message(mailMessage) # envia email
        print(f'Email da loja {loja} enviado!')
    enviarRelatorioLoja() 


#Ranking de diretorias
faturamentoLojas = vendas.groupby("Loja")[['Valor Final']].sum()
faturamentoLojasAno = faturamentoLojas.sort_values(by="Valor Final", ascending=False)

nomeArquivo = f'{diaIndicador.month}_{diaIndicador.day}_RankingAnual.xlsx'
faturamentoLojasAno.to_excel(r'Backup Arquivos Lojas\{}'.format(nomeArquivo))

vendasDia = vendas.loc[vendas['Data']==diaIndicador, :]
faturamentoLojasDia = vendasDia.groupby('Loja')[['Valor Final']].sum()
faturamentoLojasDia = faturamentoLojasDia.sort_values(by="Valor Final", ascending=False)

nomeArquivo = f'{diaIndicador.month}_{diaIndicador.day}_RankingDia.xlsx'
faturamentoLojasDia.to_excel(r'Backup Arquivos Lojas\{}'.format(nomeArquivo))


def enviarRelatorioDiretoria():      
    nome = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
    destinatarios = [nome, "seuEmail@gmail.com"]
    mailMessage = MIMEMultipart() #cria o email
    
    mailMessage["From"] = "seuEmail@gmail.com" #remetente
    mailMessage["To"] = ','.join(destinatarios) #destinatario
    
    mailMessage["Subject"] = f'Ranking Dia {diaIndicador.day}/{diaIndicador.month}' #cabeçalho
    
    # corpo do email
    
    mailBody = mailMessage.HTMLBody = f'''
    <p>Prezado, bom dia</p>

    <p>Melhor loja do Dia em Faturamento: Loja {faturamentoLojasDia.index[0]} com Faturamento R${faturamentoLojasDia.iloc[0, 0]}</p>
    <p>Pior loja do Dia em Faturamento: Loja {faturamentoLojasDia.index[-1]} com Faturamento R${faturamentoLojasDia.iloc[-1, 0]}</p>

    <p>Melhor loja do Ano em Faturamento: Loja {faturamentoLojasAno.index[0]} com Faturamento R${faturamentoLojasAno.iloc[0, 0]}</p>
    <p>Pior loja do Ano em Faturamento: Loja {faturamentoLojasAno.index[-1]} com Faturamento R${faturamentoLojasAno.iloc[-1, 0]}</p>

    <p>Segue em anexo os rankings do ano e do dia de todas as lojas.</p>
    '''
    
    mailMessage.attach(MIMEText(mailBody, "html")) 

    # anexos (pode por quantos quiser)
    attachment1 = pathlib.Path.cwd()/caminhoBackup/f'{diaIndicador.month}_{diaIndicador.day}_RankingAnual.xlsx'
    # abre o arquivo, lê e o adiciona ao email
    with open(attachment1, "rb") as arquivo:
        mailMessage.attach(MIMEApplication(arquivo.read(), name="Relatorio Ano.xlsx"))
    # Anexo 2 
    attachment2 = pathlib.Path.cwd()/caminhoBackup/f'{diaIndicador.month}_{diaIndicador.day}_RankingDia.xlsx'
    # abre o arquivo, lê e o adiciona ao email
    with open(attachment2, "rb") as arquivo:
        mailMessage.attach(MIMEApplication(arquivo.read(), name="Relatorio Dia.xlsx"))
    
    
    # conexão com o servidor
    servidor = smtplib.SMTP("smtp.gmail.com", 587)
    servidor.starttls() # formato de criptografia para enviar email em segurança
    # conexao o .env para acessar password de forma segura
    load_dotenv()
    password = os.getenv("PASSWORD")
    
    servidor.login(mailMessage["From"], password) # conecta o servidor ao seu email
    servidor.send_message(mailMessage) # envia email
    print(f'Email da Diretoria {loja} enviado!')
enviarRelatorioDiretoria()
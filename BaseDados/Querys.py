import sqlite3
import random
import string
import datetime
import os # Por causa dos caminhos do SO

#### Nome Base de dados ####
def nomeBD():
    return 'basedados.db'

#### Criar BD e Tabelas ####
def criaBD_Tabelas():
    # Liga a base de dados SQLite
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

    # Criar a tabela credenciais, caso não exista
    cursor.execute('''CREATE TABLE IF NOT EXISTS credenciais (
                        id TEXT (60) PRIMARY KEY, 
                        emailenvio TEXT (100),
                        dataenvio DATE, 
                        horaenvio TEXT (8), 
                        linkaberto BOOLEAN,
                        datalinkaberto DATE, 
                        horalinkaberto TEXT (8),
                        dadospreenchidos BOOLEAN,
                        emailpreenchido TEXT (100), 
                        passwordpreenchido TEXT (100), 
                        datapreenchido DATE, 
                        horapreenchido TEXT (8), 
                        template TEXT(50)
                    )''')
    
    # Cria a tabela templates email, caso não exista
    cursor.execute('''CREATE TABLE IF NOT EXISTS templates_email (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT (50),
        assunto TEXT (80),
        conteudo TEXT,
        seguranca INTEGER (1),
        tipodestinatario TEXT (50),
        inativo BOOLEAN
    )''')

    # Cria a tabela registo tokens obtidos, caso não exista (futura versão)
    #cursor.execute("""CREATE TABLE IF NOT EXISTS tokens (
        #id INTEGER PRIMARY KEY,
        #access_token TEXT,
        #refresh_token TEXT,
        #expires_at REAL
        #)""")

    # Confirma as mudanças
    connect.commit()

    # Fechar a ligação
    connect.close()

### Adicionar registos pré-criados (defaults) no arranque da aplicação ###
def criaRegistosDefault():

    # Le o conteúdo do ficheiro HTML
    def html(ficheiro):
        with open(ficheiro, "r", encoding="utf-8") as f:
            return f.read()
    
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

    # Conta o número de registos na tabela de templates de envio de email
    cursor.execute("SELECT count(id) FROM templates_email")
    resultado = cursor.fetchone()[0]  # Obtém o valor da contagem

    # Se estiver sem registos, então insere os registos default
    if resultado == 0:
        
        ### Registo Informações Segurança ###
        # Nome do ficheiro HTML a ser lido, dentro da pasta "TemplatesEnvioEmail"
        ficheiroHtml = os.path.join("TemplatesEnvioEmail", "InformacoesSeguranca.html")
        # Conteúdo do ficheiro HTML
        conteudoHtml = html(ficheiroHtml)

        # Cria o registo        
        cursor.execute("""
                        INSERT INTO templates_email (nome, assunto, conteudo, seguranca, tipodestinatario, inativo)
                        VALUES (?, ?, ?, ?, ?, ?)""", 
                            (
                                "Informações Segurança", #nome
                                "Foram adicionadas as informações de segurança da conta Microsoft", #assunto
                                conteudoHtml, #conteudo
                                4, #segurança
                                "Todos", #tipo destinatário
                                0 #estado ativo (0) / inativo (1)
                            )
                        )

        ### Registo Recomendação Segurança ###
        # Nome do ficheiro HTML a ser lido, dentro da pasta "TemplatesEnvioEmail"
        ficheiroHtml = os.path.join("TemplatesEnvioEmail", "RecomendacaoSeguranca.html")
        # Conteúdo do ficheiro HTML
        conteudoHtml = html(ficheiroHtml)

        # Cria o registo        
        cursor.execute("""
                        INSERT INTO templates_email (nome, assunto, conteudo, seguranca, tipodestinatario, inativo)
                        VALUES (?, ?, ?, ?, ?, ?)""", 
                            (
                                "Recomendação Segurança", #nome
                                "Recomendação para a sua segurança", #assunto
                                conteudoHtml, #conteudo
                                5, #segurança
                                "Todos", #tipo destinatário
                                0 #estado ativo (0) / inativo (1)
                            )
                        )

        ### Registo Novas Aplicações ###
        # Nome do ficheiro HTML a ser lido, dentro da pasta "TemplatesEnvioEmail"
        ficheiroHtml = os.path.join("TemplatesEnvioEmail", "NovasAplicacoes.html")
        # Conteúdo do ficheiro HTML
        conteudoHtml = html(ficheiroHtml)

        # Cria o registo        
        cursor.execute("""
                        INSERT INTO templates_email (nome, assunto, conteudo, seguranca, tipodestinatario, inativo)
                        VALUES (?, ?, ?, ?, ?, ?)""", 
                            (
                                "Novas Aplicações", #nome
                                "Novas aplicações ligadas à sua conta Microsoft", #assunto
                                conteudoHtml, #conteudo
                                5, #segurança
                                "Todos", #tipo destinatário
                                0 #estado ativo (0) / inativo (1)
                            )
                        )
        ########################################


        # Confirma a inserção
        connect.commit()

    # Fechar a ligação
    connect.close()


#### INSERT tabela Credenciais ####
# Gerar ID
def gerar_id(tamanho=50):
    # Gerar um ID aleatório com letras e números e tamanho definido, neste caso 50 caracteres
    caracteres = string.ascii_letters + string.digits
    id = ''.join(random.choices(caracteres, k=tamanho))
    return id
# Valida se ID já existe na BD
def verificar_id_existente(cursor, id):
    # Verificar se o ID já existe na base de dados
    cursor.execute("SELECT 1 FROM credenciais WHERE id = ?", (id,))
    return cursor.fetchone() is not None
# Antes de efetuar o envio do email para o destinatário, cria o registo na BD. Esse registo irá devolver o ID associado ao registo criado
def insereDadosCredenciais(emailenvio, template):
    # Liga a base de dados SQLite
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

    # Gera ID único
    id = gerar_id()

    # Verifica se o ID já existe e gera um novo (se necessário)
    while verificar_id_existente(cursor, id):
        id = gerar_id()

    # Obtem a data atual
    data_atual = datetime.datetime.now().date()

    # Obtem a hora atual e converte para string no formato HH:MM:SS
    hora_atual = datetime.datetime.now().time().strftime('%H:%M:%S')

    # Insere dados
    cursor.execute("INSERT INTO credenciais (id, emailenvio, dataenvio, horaenvio, linkaberto, datalinkaberto, horalinkaberto, dadospreenchidos, emailpreenchido, passwordpreenchido, datapreenchido, horapreenchido, template) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (id, emailenvio, data_atual, hora_atual, 0, '', '', 0, '', '', '', '', template))
    # Confirma a inserção
    connect.commit()

    # Fecha a ligação
    connect.close()

    return id # Devolve o ID, por ser necessário inserir no URL a enviar por email


#### UPDATE tabela Credenciais ####
# Atualiza o registo em como o link foi clicado, embora possa ainda não ter credenciais inseridas
def atualizar_linkclicado(id):
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

    # Verifica se o link já foi aberto para determinado id
    cursor.execute("SELECT linkaberto FROM credenciais WHERE id = ?", (id,))
    resultado = cursor.fetchone()

    if resultado:
        linkaberto = resultado[0]

        # Se o campo linkaberto já estiver preenchido, não faz o update
        if linkaberto == 1:
            #print("Erro: link aberto já está preenchido para este ID.")
            connect.close()
            return
        else: # Caso não tenha sido aberto, atualiza o registo
            # Obtem a data atual
            data_atual = datetime.datetime.now().date()

            # Obtem a hora atual e converte para string no formato HH:MM:SS
            hora_atual = datetime.datetime.now().time().strftime('%H:%M:%S')

            # Atualiza dados
            cursor.execute("UPDATE credenciais SET linkaberto = 1, datalinkaberto = ?, horalinkaberto = ? WHERE id = ?",
                        (data_atual, hora_atual, id))
            # Confirma a atualização
            connect.commit()

            # Fecha a ligação
            connect.close()
# Garantir que os dados não excedem o limite do tamanho configurado para cada coluna SQL. Caso contrário, trunca
def verificar_tamanho_maximo(dados, tamanho_maximo):
    if len(dados) > tamanho_maximo:
        return dados[:tamanho_maximo]  # Truncar os dados se exceder o limite
    return dados
# Atualiza o registo na BD
def atualizar_dados(id, email, password):
    # Liga à base de dados SQLite
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

    # Verifica se o email e password já estão preenchidos para determinado id
    cursor.execute("SELECT emailpreenchido, passwordpreenchido FROM credenciais WHERE id = ?", (id,))
    resultado = cursor.fetchone()

    if resultado: # Valida se o ID existe
        existing_email, existing_password = resultado

        # Se o email e password já estiverem preenchidos, não faz o update
        if existing_email and existing_password:
            #print("Erro: email e password já estão preenchidos para este ID.")
            connect.close()
            return
        else: # Caso não exista os dados preenchidos, atualiza o registo
            # Obtem a data atual
            data_atual = datetime.datetime.now().date()

            # Obtem a hora atual e converte para string no formato HH:MM:SS
            hora_atual = datetime.datetime.now().time().strftime('%H:%M:%S')

            email = verificar_tamanho_maximo(email, 100)  # Limita 'email' a 100 caracteres
            password = verificar_tamanho_maximo(password, 100)  # Limita 'password' a 100 caracteres

            # Atualiza dados
            cursor.execute("UPDATE credenciais SET dadospreenchidos = 1, emailpreenchido = ?, passwordpreenchido = ?, datapreenchido = ?, horapreenchido = ? WHERE id = ?",
                        (email, password, data_atual, hora_atual, id))
            # Confirma a atualização
            connect.commit()

    # Fecha a ligação
    connect.close()


#### Tabela Credenciais ####
# Consulta tabela
def select_Credenciais():
    # Liga a base de dados SQLite
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

   # Select na tabela
    cursor.execute('''SELECT 
                            emailenvio,
                            dataenvio,
                            horaenvio,
                            CASE 
                                WHEN linkaberto = 1 THEN 'Sim'
                                WHEN linkaberto = 0 THEN 'Não'
                                ELSE '-'  -- Caso o valor seja diferente de 0 ou 1
                            END AS linkaberto,
                            datalinkaberto, 
                            horalinkaberto,
                            CASE 
                                WHEN dadospreenchidos = 1 THEN 'Sim'
                                WHEN dadospreenchidos = 0 THEN 'Não'
                                ELSE '-'  -- Caso o valor seja diferente de 0 ou 1
                            END AS dadospreenchidos,
                            emailpreenchido,
                            passwordpreenchido,
                            datapreenchido,
                            horapreenchido,
                            template
                        FROM 
                            credenciais''')
    
    # Obtém os resultados da consulta
    resultados = cursor.fetchall()

    # Fecha a conexão com a base de dados
    connect.close()

    # Retorna os resultados obtidos
    return resultados


#### Templates Email ####
# Consulta Tabela
def select_TemplatesEmail():
    # Liga a base de dados SQLite
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

   # Select na tabela
    cursor.execute('''SELECT 
                            id,
                            nome,
                            assunto,
                            conteudo,
                            seguranca,
                            tipodestinatario,
                            CASE 
                                WHEN inativo = 1 THEN 'Sim'
                                WHEN inativo = 0 THEN 'Não'
                                ELSE '-'  -- Caso o valor seja diferente de 0 ou 1
                            END AS inativo
                        FROM 
                            templates_email''')
    
    # Obtém os resultados da consulta
    resultados = cursor.fetchall()

    # Fecha a conexão com a base de dados
    connect.close()

    # Retorna os resultados obtidos
    return resultados
# Utilizado para o envio de emails - surge apenas o nome do template e todos os que não estejam em status inativo
def select_TemplatesEmailEnvioEmail():
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

   # Select à tabela
    cursor.execute('''SELECT 
                            id,
                            nome,
                            assunto,
                            conteudo
                        FROM 
                            templates_email
                        WHERE
                            inativo = 0''')
    
    # Obtém os resultados
    resultados = cursor.fetchall()

    # Fecha a conexão com a base de dados
    connect.close()

    # Retorna os resultados
    return resultados
# Alterar registo
def update_dadosTemplateEmail(id, nome, assunto, conteudo, grausegurança, tipodestinatario, inativo):
    # Liga à base de dados SQLite
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

    # Atualiza dados
    cursor.execute("UPDATE templates_email SET nome = ?, assunto = ?, conteudo = ?, seguranca = ?, tipodestinatario = ?, inativo = ? WHERE id = ?",
                   (nome, assunto, conteudo, grausegurança, tipodestinatario, inativo, id))
    # Confirma a atualização
    connect.commit()

    # Fechar a ligação
    connect.close()
# Cria registo
def insert_TemplateEmail(nome, assunto, conteudo, grausegurança, tipodestinatario, inativo):
    # Liga à base de dados SQLite
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

    # Insere os dados
    cursor.execute("INSERT INTO templates_email (nome, assunto, conteudo, seguranca, tipodestinatario, inativo) VALUES (?, ?, ?, ?, ?, ?)",
                    (nome, assunto, conteudo, grausegurança, tipodestinatario, inativo))
    # Confirma a atualização
    connect.commit()

    # Fecha a ligação
    connect.close()
# Elimina registo
def delete_TemplateEmail(id):
    # Liga à base de dados SQLite
    connect = sqlite3.connect(nomeBD())
    cursor = connect.cursor()

    # Elimina dados
    cursor.execute("DELETE FROM templates_email WHERE id = ?", (id,))
    # Confirma a atualização
    connect.commit()

    # Fecha a ligação
    connect.close()


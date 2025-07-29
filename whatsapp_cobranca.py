"""
Script para envio automatizado de mensagens de cobrança via WhatsApp
Autor: [Ramon Dantas]
Data: [28/07/2025]
"""


import pywhatkit  # Biblioteca principal para enviar mensagens pelo WhatsApp
import time       # Para controlar pausas entre mensagens
from datetime import datetime  # Para trabalhar com datas e horários
import pandas as pd  # Para ler e manipular dados do arquivo Excel

# ==============================================
# CONFIGURAÇÕES INICIAIS
# ==============================================

# Carregar lista de contatos (arquivo Excel)
try:
    # Lê o arquivo Excel com os contatos
    # O arquivo deve ter colunas: nome, telefone, valor, data_vencimento
    contatos_df = pd.read_excel('teste.xlsx')
except Exception as e:
    print(f"Erro ao carregar arquivo de contatos: {str(e)}")
    exit()  # Encerra o programa se houver erro

# Template da mensagem de cobrança
# Os placeholders {nome}, {valor} e {vencimento} serão substituídos pelos dados de cada cliente
mensagem = """Olá {nome}, tudo bem?

Este é um lembrete amigável sobre o pagamento pendente no valor de R${valor} com vencimento em {vencimento}.

Por favor, regularize sua situação o quanto antes.

Atenciosamente,
Sua Empresa"""

# Tempo de espera entre mensagens (em segundos)
# Evita sobrecarregar o WhatsApp e possíveis bloqueios
intervalo = 15

# ==============================================
# PREPARAÇÃO DO HORÁRIO DE ENVIO
# ==============================================

# Define o horário inicial de envio (hora atual + 2 minutos)
# Isso dá tempo para o WhatsApp Web carregar
hora_envio = datetime.now().hour
minuto_envio = datetime.now().minute + 2

# Ajusta o horário se passar de 60 minutos
if minuto_envio >= 60:
    minuto_envio -= 60
    hora_envio += 1

# ==============================================
# LOOP PRINCIPAL DE ENVIO DE MENSAGENS
# ==============================================

print("Iniciando envio de mensagens...")

# Itera sobre cada contato no DataFrame
for index, contato in contatos_df.iterrows():
    try:
        # Personaliza a mensagem com os dados do cliente atual
        msg_personalizada = mensagem.format(
            nome=contato['nome'],
            valor=contato['valor'],
            vencimento=contato['data_vencimento']
        )
        
        # Envia a mensagem via WhatsApp
        pywhatkit.sendwhatmsg(
            phone_no=f"+55{contato['telefone']}",  # Número com código do Brasil
            message=msg_personalizada,
            time_hour=hora_envio,
            time_min=minuto_envio,
            wait_time=15  # Tempo de espera para abrir o WhatsApp Web
        )
        
        print(f"Mensagem enviada para {contato['nome']} ({contato['telefone']})")
        
        # Atualiza o horário para a próxima mensagem
        minuto_envio += 2  # Adiciona 2 minutos para o próximo envio
        if minuto_envio >= 60:
            minuto_envio -= 60
            hora_envio += 1
        
        # Espera o intervalo definido antes de enviar a próxima mensagem
        time.sleep(intervalo)
        
    except Exception as e:
        print(f"Erro ao enviar para {contato['nome']}: {str(e)}")
        # Continua para o próximo contato mesmo se der erro

print("Processo concluído!")
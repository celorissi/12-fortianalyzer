import paramiko
import time
import re

# Informações de conexão SSH
fortianalyzer_ip = '192.168.15.14'
username = 'admin'
password = 'Admin@123'

# Nome do arquivo de saída
output_file = 'output_ssh_session.txt'

# Criando uma instância do cliente SSH
ssh_client = paramiko.SSHClient()

# Ignorando o aviso de chave host desconhecida
ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

try:
    # Conectando ao FortiAnalyzer
    ssh_client.connect(fortianalyzer_ip, username=username, password=password, timeout=5)

    # Abrindo um canal SSH
    ssh_shell = ssh_client.invoke_shell()

    # Esperando um momento para que a shell esteja pronta
    time.sleep(2)

    # Limpar buffer inicial
    output = ssh_shell.recv(65535).decode('utf-8')

    # Escrevendo o início da conexão no arquivo de saída
    with open(output_file, 'w') as f:
        f.write(output)

    # Enviando o comando "diagnose dvm task repair"
    ssh_shell.send('diagnose dvm task repair\r')
    time.sleep(2)

    # Esperando a solicitação de confirmação "(y/n)"
    while not re.search(r'Do you want to continue\? \(y/n\)', output):
        output += ssh_shell.recv(65535).decode('utf-8')
        time.sleep(1)

    # Imprimindo a solicitação de confirmação
    print(output)

    # Enviando a resposta "y" para continuar
    ssh_shell.send('y\r\n')
    time.sleep(2)

    # Lendo a saída após enviar "y"
    output = ''
    while True:
        # Aguardando até que a saída contenha o prompt de comando novamente
        if ssh_shell.recv_ready():
            data = ssh_shell.recv(65535)
            output += data.decode('utf-8')
            with open(output_file, 'a') as f:
                f.write(data.decode('utf-8'))
            if output.endswith('# '):
                break
        else:
            time.sleep(1)

    # Imprimindo a saída final após o comando
    print(output)

finally:
    # Fechando a conexão SSH
    ssh_client.close()
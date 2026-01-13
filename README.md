# Painel NDI Free Audio

O Painel NDI Free Audio é uma interface gráfica (GUI) desenvolvida em Python para simplificar o gerenciamento do NDI Free Audio 6.2.1.0. Ele permite transformar dispositivos de som locais em fontes NDI (Input) ou receber áudio de outras fontes NDI na rede (Output), tudo isso rodando como um serviço do Windows para garantir estabilidade e inicialização automática.

## 🚀 Funcionalidades
Interface Amigável: Criada com customtkinter para um visual moderno e escuro.
Gestão de Serviços: Instale, reinicie, pare ou exclua fluxos de áudio como serviços do Windows (via NSSM).
Scanner de Dispositivos: Lista automaticamente os IDs de áudio do sistema para facilitar a configuração.
Modos Flexíveis: * ENTRADA: Envia o som do PC local para a rede via NDI.
SAÍDA: Recebe som de um PC remoto via NDI e reproduz no dispositivo local.
Utilitários Integrados: * Instalador do VB-CABLE incluído.
Configuração automática de regras no Firewall do Windows.
Controle de Ganho (dB) nativo.

## 🛠️ Pré-requisitos
Para que o painel funcione corretamente, a estrutura de pastas deve conter:
nssm.exe (na pasta raiz do app).
NDIFreeAudio.exe (na pasta raiz do app).
Uma pasta chamada VBCABLE contendo o VBCABLE_Setup_x64.exe.

Importante: O aplicativo deve ser executado como Administrador para conseguir gerenciar serviços e regras de firewall.

## 📥 Instalação de Dependências (Apps de Terceiros)

NDI Free Audio:
1. Baixe em: ndi.video/tools/free-audio
2. Copie o arquivo executável para a pasta "C:/Painel NDI Free Audio/"
Renomeie o arquivo para "NDIFreeAudio.exe"

NSSM (Non-Sucking Service Manager):
1. Baixe em: nssm.cc/download
2. Copie o executável para "C:/Painel NDI Free Audio/"
3. Renomeie o arquivo para "nssm.exe"

VB-CABLE Driver:
1. Baixe o arquivo VBCABLE_Driver_Pack45.zip em vb-audio.com/Cable
2. Extraia todo o conteúdo do arquivo .zip para a pasta: "C:/Painel NDI Free Audio/VBCABLE/"

## 📖 Como Usar
1. Identificar Dispositivos
Clique em "ESCANEAR DISPOSITIVOS" na coluna da direita. O log mostrará os IDs disponíveis (ex: 0, 1, 2). Anote o ID do microfone ou da saída de som que deseja usar.
2. Criar um Fluxo de Entrada (PC Local -> Rede)
Selecione o modo ENTRADA.
Insira o ID do Dispositivo (obtido no passo anterior).
Dê um Nome ao Stream (ex: MesaSom_Igreja).
(Opcional) Ajuste o Ganho em dB.
Clique em CRIAR E INICIAR FLUXO.
3. Criar um Fluxo de Saída (Rede -> Som Local)
Selecione o modo SAÍDA.
Insira o ID do Dispositivo onde o som deve sair.
Insira o Nome do PC de Origem (Host) que está transmitindo o NDI.
Insira o Nome do Stream original.
Clique em CRIAR E INICIAR FLUXO.

4. Gestão de Serviços
Na parte superior, você verá uma lista de serviços ativos. Você pode reiniciar ou remover fluxos antigos para manter seu sistema limpo.

## 🔧 Solução de Problemas
O áudio não chega? Clique no botão "LIBERAR FIREWALL". O app criará regras automáticas de entrada e saída para o executável de áudio.
Precisa de cabos virtuais? Use o botão "INSTALAR VB-CABLE" para abrir o instalador do driver de áudio virtual.
Serviço não inicia? Verifique se o ID do dispositivo está correto e se o NDIFreeAudio.exe está na mesma pasta do painel.

## 💻 Tecnologias Utilizadas
Python 3, CustomTkinter, NSSM (Non-Sucking Service Manager), PowerShell (para automação de comandos de sistema).

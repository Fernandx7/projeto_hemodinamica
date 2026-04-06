#!/bin/bash

SERVICE_NAME="laudo_ia.service"
SERVICE_PATH="/home/fernando/projeto_hemodinamica/$SERVICE_NAME"

echo "Instalando serviço $SERVICE_NAME..."

# Copiar para a pasta do systemd
sudo cp "$SERVICE_PATH" /etc/systemd/system/

# Recarregar daemon e habilitar serviço
sudo systemctl daemon-reload
sudo systemctl enable "$SERVICE_NAME"
sudo systemctl restart "$SERVICE_NAME"

echo "Status do serviço:"
sudo systemctl status "$SERVICE_NAME"

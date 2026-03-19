#!/bin/bash
echo "Compilando AppHemo de testes..."
javac -cp "bin/AppHemo.jar" testes/AppHemo.java

if [ $? -eq 0 ]; then
    echo "Compilação concluída com sucesso!"
    echo "Para rodar use: java -cp \"bin/AppHemo.jar:testes\" AppHemo"
else
    echo "Erro na compilação. Verifique se o 'javac' está instalado."
fi

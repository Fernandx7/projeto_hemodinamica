# 🏥 Guia Técnico Profundo - Ecossistema Hemodinâmica (VOTICOR)

Este manual detalha a arquitetura, as decisões de engenharia e o fluxo operacional do sistema de automação hospitalar voltado para Cateterismo e Angioplastia.

---

## 🏗 Arquitetura do Sistema

O sistema opera em um modelo de **microserviços coordenados**, dividindo as responsabilidades para garantir que falhas em um módulo (ex: IA de voz) não parem o fluxo hospitalar principal (recepção e sala).

1.  **Backend Central (Flask - 2400)**: O "Cérebro". Gerencia persistência, arquivos e coordena a transição de estados do paciente.
2.  **Módulo de Percepção (FastAPI - 8000)**: O "Ouvido". Especializado em processamento de áudio pesado e integração com LLM (Gemini).
3.  **Terminal Médico (Java Swing)**: O "Console de Controle". Interface robusta para uso em desktops médicos com foco em agilidade.

---

## 📂 Detalhamento de Componentes (Pasta `atualização mega/`)

### 1. `site.py` (O Integrador)
Este script é o motor que transforma uma ficha administrativa em um laudo médico técnico.
*   **Quarentena Automática**: Ao receber um arquivo, ele o isola na pasta `aguardando_procedimento`. Isso impede que o médico veja um laudo que ainda não tem os materiais da sala.
*   **Extração Híbrida**: Utiliza o modelo **Gemini 3.1 Flash Lite** para ler a ficha original. Se a IA falhar, um **Regex de 3 camadas** entra em ação para extrair o Nome, CNS e Nascimento sem confundir com a procedência.
*   **Motor de Injeção**: No método `finalizar_mesa`, ele abre o modelo `.docx`, localiza a tag `{{materiais}}` e injeta a lista vinda do tablet, formatando-a para evitar o bug de parágrafo justificado do Word.

### 2. `AppHemo.java` (O Dashboard)
Refatorado para a versão 11.0+, ele abandonou o sistema de abas por um **Painel de Controle em 3 Colunas**:
*   **Coluna 1 (Em Sala)**: Monitora o `ativo.json`. Permite ao médico ver quem está na mesa cirúrgica agora.
*   **Coluna 2 (Aguardando)**: A lista de trabalho. Implementa o **Interceptador de Modo** (Ditar ou Digitar).
*   **Técnica de Substituição "Tanque"**: Devido à fragmentação de XML do Word (onde `{{NOME}}` pode ser salvo como 3 objetos diferentes), o Java agora lê o parágrafo inteiro, limpa o XML sujo e reconstrói o texto preservando a **Fonte Arial 11** e estilos de negrito.

### 3. `teste_modelos/` (Isolamento de Produção)
Contém os modelos `.docx` que o médico realmente usa.
*   **Modelos Base**: Possuem o texto padrão completo (usado no modo "Digitar").
*   **Modelos _ditado**: Contêm apenas os títulos das seções. A IA os utiliza para que não haja sobreposição de informações antigas com o novo ditado.

---

## 📂 Detalhamento do Módulo `laudo_falado/`

### `app.py` (Processamento de Linguagem Natural)
*   **Pipeline de Áudio**: Recebe o áudio (WebM/WAV), converte para MP3 via **FFmpeg** para reduzir o consumo de banda e garantir compatibilidade com o Google Cloud.
*   **Prompt Engineering**: A IA é instruída como um cardiologista. Ela sabe que não deve repetir o cabeçalho e deve focar apenas nos achados angiográficos.
*   **Limpeza de Seção**: O script localiza marcadores como `CORONARIOGRAFIA` e apaga todo o texto abaixo dele antes de inserir a transcrição. Isso resolve o problema de laudos que ficavam com "duas conclusões".

---

## 📡 Protocolos de Comunicação (Rotas Detalhadas)

### API Flask (Porta 2400)
*   `POST /`: Recepção envia a ficha original.
*   `GET /api/sala/espera`: O Tablet consulta quem está na quarentena para iniciar.
*   `POST /api/sala/iniciar`: O Tablet notifica que o paciente entrou em mesa. Gera o `ativo.json`.
*   `POST /api/sala/atualizar`: O Tablet envia cada Stent/Balão gasto. O servidor salva no JSON.
*   `POST /api/sala/finalizar`: O Tablet encerra a cirurgia. O servidor gera os laudos base, injeta materiais e move para `uploads/`.
*   `GET /api/pendentes`: O AppHemo consulta para mostrar na lista do médico.

### API de IA (Porta 8000)
*   `POST /process`: O AppHemo envia o áudio. A IA processa e retorna o JSON com o link do novo `.docx` gerado.

---

## 🛠 Padrões Técnicos e Regras de Ouro

### 1. Nomenclatura e Pastas
*   **CAT**: Pasta simples `[NUM_EXAME] [NOME]`.
*   **PTCA (Angioplastia)**: Pasta e arquivo DEVEM conter o sufixo ` PTCA`. O sistema usa esse sufixo para decidir qual modelo de Word carregar.

### 2. Formatação de Texto
*   Todo o corpo clínico deve ser **Arial 11**.
*   **PascalCase**: O sistema limpa nomes em caixa alta. `JOÃO DA SILVA` -> `João da Silva`. Isso é aplicado no Java no momento final da geração.

### 3. Controles de Conteúdo (SDT)
O sistema busca por IDs específicos no XML do Word para preencher dados sem usar tags visíveis:
*   `campo_nome`, `campo_cns`, `campo_nasc`, `campo_procedencia`, `campo_num_exame`, `campo_materiais`.

---
*Este guia deve ser atualizado sempre que uma nova lógica de preenchimento ou rota for adicionada.*

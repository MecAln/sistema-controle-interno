# 💻 Sistema de Controle Interno e Automação de RMA

![Status](https://img.shields.io/badge/Status-Concluído-success)
![Tecnologias](https://img.shields.io/badge/Stack-JavaScript_|_Google_Apps_Script_|_HTML_|_CSS-blue)

Um sistema web completo (Single Page Application) desenvolvido para centralizar, automatizar e gerenciar as operações diárias dentro do setor de Atendimento na parte de Devoluções e Cancelamentos.

## 🎯 O Problema
Antes do sistema, o controle de devoluções, solicitações de reembolso, bloqueios logísticos e a padronização de procedimentos eram feitos de forma descentralizada. Isso gerava retrabalho, falhas de comunicação com transportadoras, dificuldade na extração de relatórios e falta de rastreabilidade das ações dos analistas.

## 💡 A Solução
Desenvolvi uma aplicação web utilizando o ecossistema do **Google Workspace**. O frontend foi construído em HTML/CSS e JavaScript puro (Vanilla), enquanto o backend roda via **Google Apps Script**, utilizando planilhas do **Google Sheets** como banco de dados e o **Google Drive** para armazenamento de arquivos e imagens.

## 🚀 Principais Funcionalidades

* 🔐 **Autenticação e Permissões:** Sistema de login com criptografia básica e controle de acesso baseado em funções (Admin vs. Analista). Telas dinâmicas renderizadas de acordo com a permissão do usuário.
* 💸 **Gestão de Reembolsos:** Formulários complexos que processam dados financeiros e disparam **e-mails automatizados** para os setores responsáveis, já com os arquivos (NFs) anexados em Base64/Blob.
* 🚚 **Bloqueio Logístico (Transportadoras & PLP):** Módulo de integração que envia solicitações automáticas de barragem de pedidos diretamente para as transportadoras configuradas no sistema.
* 📚 **Base de Conhecimento (Procedimentos):** Um portal interno para os analistas consultarem regras. Utiliza o **Quill.js** como editor Rich Text, convertendo imagens coladas para base64 e salvando-as automaticamente na nuvem (Google Drive).
* 📊 **Relatórios Inteligentes:** Geração de dashboards e gráficos interativos em tempo real utilizando a biblioteca **Google Charts**, com opção de consolidar os dados e exportar uma planilha oficial de fechamento mensal.
* 👁️ **Auditoria (Logs):** Registro automático e invisível no backend de todas as ações sensíveis (criação, edição e exclusão) realizadas por cada usuário logado.

## 🛠️ Stack Tecnológica

* **Frontend:** HTML5, CSS3 (Design responsivo, Dark Mode, animações CSS) e JavaScript.
* **Backend:** Google Apps Script (V8).
* **Banco de Dados & Storage:** Google Sheets e Google Drive API.
* **Bibliotecas Adicionais:** 
  * `Google Charts` (Dashboards)
  * `Quill.js` (Editor de Texto)
  * `Tom Select` (Dropdowns interativos)
  * `Simple-Datatables` (Tabelas de dados)

## 📸 Demonstração da Interface
> **Nota:** Por motivos de segurança e confidencialidade, partes do código contendo dados sensíveis da empresa original foram anonimizadas ou removidas. O GIF abaixo utiliza dados fictícios.

LOGIN

<img width="1375" height="716" alt="image" src="https://github.com/user-attachments/assets/85a97a0b-c0b3-4fb9-b910-1012ff5fb343" />

SESSÃO DE SOLICITAÇÃO DE REEMBOLSO

<img width="1029" height="912" alt="image" src="https://github.com/user-attachments/assets/baf28d1f-50d8-4781-a2b6-390dd3cd4db1" />

SESSÃO DOS REVERTIDOS

<img width="973" height="828" alt="image" src="https://github.com/user-attachments/assets/81a355cf-19de-4009-abe3-283a8a083381" />

SESSÃO DAS SOLICITAÇÕES DE BARRAGEM AS TRANSPORTADORAS

<img width="1114" height="897" alt="image" src="https://github.com/user-attachments/assets/28d97a31-1ffd-4359-9b52-ff042fbb2f0f" />

SESSÃO DE BARRAGEM NO CD

<img width="969" height="727" alt="image" src="https://github.com/user-attachments/assets/9d75adf2-1b6b-4fd1-9d76-efb47a22f8ac" />

SESSÃO DE BASE DE PROCEDIMENTOS DO SETOR E ANOTAÇÕES SOBRE DUVIDAS FREQUENTES

<img width="996" height="362" alt="image" src="https://github.com/user-attachments/assets/7f3f3fdc-72f8-42f3-8221-3071390998f7" />

SESSÃO DOS RELATORIOS

<img width="1026" height="863" alt="image" src="https://github.com/user-attachments/assets/03acaf27-2407-47cf-a62d-5cad174bf473" />

SESSÃO DE ADM PARA CONTROLE DE USUARIOS E TRANSPORTADORAS

<img width="1004" height="554" alt="image" src="https://github.com/user-attachments/assets/66adac24-787f-47ff-ac65-ae424feecb5e" />


## 🧠 Desafios Técnicos Superados
1. **Upload de Arquivos:** Lidar com limites de tempo de execução do Google Scripts ao fazer upload de arquivos grandes convertendo-os em Base64 e reconstruindo-os no Drive.
2. **Performance (SPA):** Estruturar o sistema em abas dinâmicas no frontend para evitar recarregamentos de página (Single Page Application), melhorando drasticamente a experiência do usuário.
3. **Segurança de Dados:** Implementar travas no backend (GS) para garantir que analistas não consigam deletar ou editar registros que não os pertencem.

---
Desenvolvido por **Allan Christian** 

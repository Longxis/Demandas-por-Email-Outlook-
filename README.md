# 🤖 Robô de Monitoramento de Relatórios (Excel + Outlook)

Automação em Python desenvolvida para **monitorar relatórios controlados em Excel** e **enviar notificações automáticas por e-mail via Outlook** sempre que houver mudanças relevantes de status.

O robô consolida todas as atualizações em **um único e-mail**, evitando notificações duplicadas e garantindo comunicação clara para o time.

---

## 📌 Funcionalidade principal

- Leitura periódica de uma planilha Excel
- Identificação de mudanças de status por relatório
- Classificação automática (finalizado, atraso, manutenção, etc.)
- Controle de memória para evitar e-mails repetidos
- Envio de e-mail HTML profissional via Outlook

---

## 📊 Base de dados (Excel)

O sistema utiliza o arquivo **`Tarefas.xlsm`** como base de controle.

### Aba principal: `Agenda`

Cada linha representa um relatório ou rotina monitorada.

Colunas utilizadas pelo robô:

- **ATIVO**: define se o relatório será monitorado (`SIM` / `NÃO`)
- **RELATORIOS**: nome do relatório (identificador único)
- **STATUS**: situação atual da rotina
- **FEITO?**: indica conclusão ou atraso
- **OBS**: observações e justificativas

Somente relatórios com `ATIVO = SIM` entram no monitoramento.

---

## 🔎 Regras de interpretação

| Situação | Condição no Excel |
|--------|------------------|
| ✅ Finalizado | STATUS = FEITO |
| 🛠️ Em manutenção | STATUS = MANUTENÇÃO e FEITO? = NÃO |
| ⏰ Atraso sem atualização | STATUS = ATRASO BASE e FEITO? = ATRASO |
| 🔄 Atraso em atualização | STATUS = ATRASO BASE e FEITO? ≠ ATRASO |
| 📅 Atualização hoje | STATUS = ATT HOJE |

---

## 📨 Envio de e-mail

- Enviado via **Outlook Desktop**
- Um único e-mail por ciclo
- Assunto padrão: **BI Abastecimento - Rotinas**
- Conteúdo em HTML com seções organizadas:
  - Atualizações previstas para hoje
  - Relatórios finalizados
  - Em manutenção
  - Atrasos (com e sem atualização prevista)
- Observações exibidas quando preenchidas no Excel

---

## 🧠 Controle de memória

O robô utiliza o arquivo:

monitor_relatorios_state.json

Esse arquivo registra:
- Último status notificado por relatório
- Data e hora do envio

Há um **TTL configurável**, evitando notificações repetidas dentro do período definido.

---

## ⏱️ Agendamento

- Execução em loop contínuo
- Intervalo configurável (padrão: 15 minutos)
- Caso não haja mudanças de status, **nenhum e-mail é enviado**

---

## ⚙️ Configurações principais

No código é possível configurar:

- Caminho do arquivo Excel
- Nome da aba
- Colunas monitoradas
- Intervalo de execução
- TTL da memória
- Destinatários do e-mail (To / CC / BCC)
- Modo de teste (visualizar e-mail antes de enviar)

---

## 🛠️ Requisitos

- Windows
- Outlook Desktop instalado e configurado
- Python 3

### Bibliotecas necessárias
```bash
pip install pandas openpyxl pywin32

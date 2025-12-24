# Local Assistant Support Notes

Este arquivo é um espaço local, não versionado, para eu (o assistente) guardar notas e instruções que ajudam no pair-coding com você.

Objetivo
- Registrar decisões de curto prazo, contextos úteis e sugestões rápidas que não devem ir para o repositório público.

Conteúdo sugerido
- Atalhos de execução locais (ex.: como ativar `.venv` e rodar `gerar_planilha.py`).
- Trechos de comandos de depuração usados frequentemente.
- Observações contextuais: por exemplo, "não forçar push sem backup".
- Pequenas notas de UX/ajustes que o usuário pediu e que ainda não virou ticket.

Exemplo de entradas rápidas
- 2025-12-24: Use `data_formatada = datetime.strftime('%d/%m/%Y')` para exibição em CSV; preserve `data_arquivo = ...strftime('%d.%m.%Y')` para nomes de arquivo no Windows.
- 2025-12-24: Ao salvar Excel, tentar remover o arquivo existente; caso falhe por arquivo em uso, salvar como ` (novo).xlsx`.

Privacidade e segurança
- Não inclua segredos, chaves de API, senhas ou informações sensíveis neste arquivo.

Como usar
- Eu atualizo este arquivo conforme trabalhamos. Você pode abri-lo localmente para ver o resumo das decisões recentes.

---

(Se desejar, posso adaptar o formato, por exemplo: JSON estruturado, TOML, ou um arquivo separado por data.)

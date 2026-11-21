#!/bin/bash
# Publica as ferramentas no GitHub Pages.
#
# Copia os arquivos de CLAUDE/meta-monitor-integracao para este clone, commita
# e envia. Só mexe no que mudou: se nada mudou, não cria commit vazio.
#
# Autenticação vem do GitHub CLI (gh auth login), guardada no keychain do Mac.
# Nenhum token fica escrito aqui.
#
# Uso:  ./publicar.sh "mensagem do commit"
set -e

ORIGEM="$HOME/CLAUDE/meta-monitor-integracao"
DESTINO="$HOME/CLAUDE/meta-ads-automation"

# Só estes sobem. Arquivo que não estiver na lista fica de fora de propósito:
# evita publicar rascunho ou dado interno sem querer.
ARQUIVOS=(
  ar.html
  erevanton-quiz.html
  meta-ads-otimizacao-arnutra.html
  meta-ads-otimizacao.html
  meta-ads-automation.html
  meta-ads-catalogos.html
  meta-ads-monitor.html
  index.html
  briefing-catalogo.html
)

cd "$DESTINO"
# --autostash: o proprio script pode ter sido editado antes de rodar, e sem
# isso o pull morre com "you have unstaged changes" e segue adiante sem
# sincronizar, que e pior do que falhar.
git pull --quiet --rebase --autostash origin main || true

for f in "${ARQUIVOS[@]}"; do
  if [ -f "$ORIGEM/$f" ]; then
    cp "$ORIGEM/$f" "$DESTINO/$f"
  fi
done

# add antes de olhar o que mudou: `git diff` sozinho nao enxerga arquivo novo,
# e a primeira publicacao do ar.html passou sem aparecer na lista.
git add -A

if git diff --cached --quiet; then
  echo "Nada mudou. Nenhum commit criado."
  exit 0
fi

echo "Mudou:"
git diff --cached --stat | sed 's/^/  /'
git commit --quiet -m "${1:-Atualiza ferramentas}"
git push --quiet origin main
echo "Publicado. O GitHub Pages leva 1 a 2 minutos para refletir."

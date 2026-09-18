#!/usr/bin/env bash
# vybe-test: bash/append_redirection/rebind
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
f="${TMPDIR:-/tmp}/bash_append_redirection_${BASHPID}_${16}"
: > "$f"
printf 'start\n' > "$f"
if (( 16 % 2 == 0 )); then
  for ((i=0; i<16; i++)); do
    printf '%s\n' "$i" >> "$f"
  done
else
  { for ((i=0; i<16; i++)); do
      printf '%s\n' "$i"
    done
  } >> "$f"
fi
count=0
while IFS= read -r _; do
  count=$((count + 1))
done < "$f"
expected=$((16 + 1))
[ "$count" -eq "$expected" ] || fail "append redirection count mismatch"
rm -f "$f"
echo PASS
exit 0

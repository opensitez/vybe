#!/usr/bin/env bash
# vybe-test: bash/reserved_words_and_keywords/esac_as_case_pattern_needs_open_paren
# Right after "in", the word esac closes the case; a leading ( lets it be a pattern.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
out=$(case esac in (esac) echo matched;; esac)
[ "$out" = matched ] || fail "want [matched] got [$out]"
eval 'case esac in esac) echo bad;; esac' 2>/dev/null; st=$?
[ "$st" -ne 0 ] || fail "bare esac pattern must be a syntax error"
echo PASS
exit 0

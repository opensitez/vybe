<?php
// vybe-test: php/mb_strings/mb_word_wrap
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

function mb_word_count(string $s): int {
    return count(array_filter(preg_split('/\s+/u', $s), fn($w) => $w !== ''));
}
echo mb_word_count("Hello World PHP");
echo mb_word_count("  spaces  everywhere  ");

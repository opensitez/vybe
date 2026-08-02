<?php
// vybe-test: php/mb_strings/mb_string_reverse
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

function mb_strrev(string $s): string {
    return implode('', array_reverse(mb_str_split($s)));
}
echo mb_strrev("hello");
echo mb_strrev("日本語");

<?php
// vybe-test: php/mb_strings/mb_truncate_string
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

function mb_truncate(string $s, int $maxLen, string $suffix = '...'): string {
    if (mb_strlen($s) <= $maxLen) return $s;
    return mb_substr($s, 0, $maxLen - mb_strlen($suffix)) . $suffix;
}
echo mb_truncate("Hello World", 8);
echo mb_truncate("Hi", 10);

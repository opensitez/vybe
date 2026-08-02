<?php
// vybe-test: php/string_extra_builtins/strtok_tokenize_by_delimiter
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$token = strtok("Hello World PHP", " ");
$parts = [];
while ($token !== false) {
    $parts[] = $token;
    $token = strtok(" ");
}
echo count($parts);
echo $parts[0];

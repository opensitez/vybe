<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_grep_array_filtering
// origin: languages/php/tests/php/test_php_regular_expressions_pcre_matching.rs
// vybe-test-mode: compile

$input = ["123", "abc", "456", "def"];
$numbers = preg_grep('/^\d+$/', $input);
echo implode(",", $numbers);

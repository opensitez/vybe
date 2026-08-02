<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_replace_array_patterns_replacements
// origin: languages/php/tests/php/test_php_regular_expressions_pcre_matching.rs
// vybe-test-mode: compile

$patterns = ['/quick/', '/brown/', '/fox/'];
$replacements = ['bear', 'black', 'wolf'];
echo preg_replace($patterns, $replacements, 'The quick brown fox jumps');

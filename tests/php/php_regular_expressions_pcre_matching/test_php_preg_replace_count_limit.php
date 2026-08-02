<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_replace_count_limit
// origin: languages/php/tests/php/test_php_regular_expressions_pcre_matching.rs
// vybe-test-mode: compile

$count = 0;
$result = preg_replace('/foo/', 'bar', 'foo foo foo', limit: 2, count: $count);
echo "Result=$result Count=$count";

<?php
// vybe-test: php/pcre_advanced/preg_replace_backreference
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$result = preg_replace('/(\w+)\s+(\w+)/', '$2 $1', 'Hello World');
echo $result;
$result2 = preg_replace('/(\d{4})-(\d{2})-(\d{2})/', '$3/$2/$1', '2024-06-15');
echo $result2;

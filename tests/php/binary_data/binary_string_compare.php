<?php
// vybe-test: php/binary_data/binary_string_compare
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$a = "\x01\x02\x03";
$b = "\x01\x02\x03";
$c = "\x01\x02\x04";
echo ($a === $b) ? 'equal' : 'not equal';
echo ($a === $c) ? 'equal' : ':not equal';

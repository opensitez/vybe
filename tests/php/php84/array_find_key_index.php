<?php
// vybe-test: php/php84/array_find_key_index
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

$scores = [45, 78, 92, 61, 88];
$idx = array_find_key($scores, fn($s) => $s >= 90);
echo $idx;

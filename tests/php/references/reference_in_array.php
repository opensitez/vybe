<?php
// vybe-test: php/references/reference_in_array
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$a = 1;
$arr = [&$a, 2, 3];
$a = 100;
echo $arr[0];

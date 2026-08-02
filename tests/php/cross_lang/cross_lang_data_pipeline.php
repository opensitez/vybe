<?php
// vybe-test: php/cross_lang/cross_lang_data_pipeline
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

// PHP array operations produce same bytecode as Python/JS equivalents
$data = range(1, 10);
$doubled = array_map(fn($n) => $n * 2, $data);
$evens = array_filter($doubled, fn($n) => $n % 4 == 0);
$sum = array_reduce($evens, fn($c, $i) => $c + $i, 0);
echo $sum;

<?php
// vybe-test: php/php5_legacy/array_pipeline
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

$data = ['banana', 'apple', 'cherry', 'date'];
sort($data);
$upper = array_map(fn($s) => strtoupper($s), $data);
$filtered = array_filter($upper, fn($s) => strlen($s) > 4);
echo implode(', ', $filtered);

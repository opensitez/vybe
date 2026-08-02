<?php
// vybe-test: php/match_advanced/match_chained_transformation
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$input = 'EUR';
$symbol = match($input) { 'USD' => '$', 'EUR' => '€', 'GBP' => '£', default => '?' };
$rate   = match($input) { 'USD' => 1.0, 'EUR' => 1.08, 'GBP' => 1.27, default => 0.0 };
echo "$symbol" . number_format($rate, 2);

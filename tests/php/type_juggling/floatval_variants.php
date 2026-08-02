<?php
// vybe-test: php/type_juggling/floatval_variants
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

echo floatval('1.5e3') . "\n";   // 1500
echo floatval('  -2.5  ') . "\n"; // -2.5
echo doubleval('3.14') . "\n";

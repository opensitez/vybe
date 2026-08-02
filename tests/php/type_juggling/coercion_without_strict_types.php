<?php
// vybe-test: php/type_juggling/coercion_without_strict_types
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

// Without strict_types, PHP coerces args
function addNums(int $a, int $b): int { return $a + $b; }
echo addNums("3", "4");  // coerces strings to ints
echo addNums(2.9, 1.1);  // coerces floats to ints (truncates)

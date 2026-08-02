<?php
// vybe-test: php/type_juggling/spaceship_mixed_types
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

echo (1 <=> 2)   . "\n"; // -1
echo (2 <=> 2)   . "\n"; // 0
echo (3 <=> 2)   . "\n"; // 1
echo ("a" <=> "b") . "\n"; // -1
echo ([1,2] <=> [1,2]) . "\n"; // 0
echo ([1,3] <=> [1,2]) . "\n"; // 1

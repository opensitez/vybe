<?php
// vybe-test: php/functional_style/usort_spaceship_operator
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

$words = ['banana', 'apple', 'cherry', 'date'];
usort($words, fn($a, $b) => $a <=> $b);
echo implode(',', $words);

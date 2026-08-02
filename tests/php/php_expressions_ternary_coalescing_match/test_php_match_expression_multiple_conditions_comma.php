<?php
// vybe-test: php/php_expressions_ternary_coalescing_match/test_php_match_expression_multiple_conditions_comma
// origin: languages/php/tests/php/test_php_expressions_ternary_coalescing_match.rs
// vybe-test-mode: compile

$char = "e";
$isVowel = match (strtolower($char)) {
    "a", "e", "i", "o", "u" => true,
    default => false,
};
echo $isVowel ? "VOWEL" : "CONSONANT";

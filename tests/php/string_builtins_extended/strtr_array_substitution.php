<?php
// vybe-test: php/string_builtins_extended/strtr_array_substitution
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$map = ["apple" => "fruit", "dog" => "animal", "blue" => "color"];
$result = strtr("apple and dog and blue", $map);
echo $result;

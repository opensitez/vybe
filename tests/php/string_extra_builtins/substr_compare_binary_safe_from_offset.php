<?php
// vybe-test: php/string_extra_builtins/substr_compare_binary_safe_from_offset
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$result = substr_compare("abcdefg", "cde", 2, 3);
echo $result === 0 ? "equal" : "not-equal";
$diff = substr_compare("abcdefg", "xyz", 0, 3);
echo $diff !== 0 ? "different" : "same";

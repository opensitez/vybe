<?php
// vybe-test: php/string_extra_builtins/strpbrk_search_for_any_char_in_set
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$result = strpbrk("This is a test", "aeiou");
echo is_string($result) ? "found" : "not";
$none = strpbrk("bcdfg", "aeiou");
echo $none === false ? "false" : "found";

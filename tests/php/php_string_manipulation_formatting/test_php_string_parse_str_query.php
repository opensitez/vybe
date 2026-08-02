<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_parse_str_query
// origin: languages/php/tests/php/test_php_string_manipulation_formatting.rs
// vybe-test-mode: compile

$str = "first=value&arr[]=foo+bar&arr[]=baz";
parse_str($str, $output);
echo $output['first'];

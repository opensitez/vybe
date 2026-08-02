<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_key_first_and_last_php73
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs
// vybe-test-mode: compile

$a = ["first" => 1, "mid" => 2, "last" => 3];
echo "First=" . array_key_first($a) . " Last=" . array_key_last($a);

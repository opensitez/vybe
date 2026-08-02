<?php
// vybe-test: php/string_builtins_extended/str_getcsv_parse_csv_line
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$fields = str_getcsv("one,two,three");
echo count($fields);
echo $fields[0];
echo $fields[2];

<?php
// vybe-test: php/string_builtins_extended/str_getcsv_quoted_fields_and_escapes
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$fields = str_getcsv('"a,b","c\"d","e\"f"');
echo $fields[0];
echo $fields[1];
echo $fields[2];

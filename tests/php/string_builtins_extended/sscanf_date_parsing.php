<?php
// vybe-test: php/string_builtins_extended/sscanf_date_parsing
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$parts = sscanf("2024-07-15", "%d-%d-%d");
echo $parts[0];
echo $parts[1];
echo $parts[2];

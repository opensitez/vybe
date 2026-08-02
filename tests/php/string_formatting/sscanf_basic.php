<?php
// vybe-test: php/string_formatting/sscanf_basic
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$result = sscanf("Age: 25", "Age: %d");
echo $result[0];
echo "\n";

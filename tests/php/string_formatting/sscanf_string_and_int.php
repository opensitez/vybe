<?php
// vybe-test: php/string_formatting/sscanf_string_and_int
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

[$name, $age] = sscanf("Alice 30", "%s %d");
echo "$name is $age";
echo "\n";

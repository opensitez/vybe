<?php
// vybe-test: php/string_formatting/sscanf_multiple
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

[$y, $m, $d] = sscanf("2024-01-15", "%d-%d-%d");
echo "$y-$m-$d";
echo "\n";

<?php
// vybe-test: php/php_compact_extract_superglobals/test_php_extract_skip_existing_variables
// origin: languages/php/tests/php/test_php_compact_extract_superglobals.rs
// vybe-test-mode: compile

$status = "protected";
$input = ["status" => "overwritten", "new_key" => "value"];

extract($input, EXTR_SKIP);
echo "$status | $new_key";

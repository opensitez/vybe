<?php
// vybe-test: php/php_compact_extract_superglobals/test_php_extract_if_exists_import
// origin: languages/php/tests/php/test_php_compact_extract_superglobals.rs
// vybe-test-mode: compile

$existing = "initial";
$input = ["existing" => "updated", "non_existing" => "ignored"];

extract($input, EXTR_IF_EXISTS);
echo "$existing " . (isset($non_existing) ? "YES" : "NO");

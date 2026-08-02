<?php
// vybe-test: php/php_mbstring_multibyte_utf8_ops/test_php_mb_scrub_invalid_encoding
// origin: languages/php/tests/php/test_php_mbstring_multibyte_utf8_ops.rs
// vybe-test-mode: compile

$invalid = "Hello \xFF\xFE World";
$scrubbed = mb_scrub($invalid, "UTF-8");
echo strlen($scrubbed) > 0 ? "SCRUBBED" : "EMPTY";

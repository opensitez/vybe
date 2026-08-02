<?php
// vybe-test: php/php_json_encode_decode_options/test_php_json_encode_hex_tag_amp_apos_quot_flags
// origin: languages/php/tests/php/test_php_json_encode_decode_options.rs
// vybe-test-mode: compile

$html = '<a href="test.php?a=1&b=2">O\'Reilly</a>';
$encoded = json_encode($html, JSON_HEX_TAG | JSON_HEX_AMP | JSON_HEX_APOS | JSON_HEX_QUOT);
echo $encoded;

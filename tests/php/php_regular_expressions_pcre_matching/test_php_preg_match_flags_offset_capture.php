<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_match_flags_offset_capture
// origin: languages/php/tests/php/test_php_regular_expressions_pcre_matching.rs
// vybe-test-mode: compile

$str = "abc 123 def";
preg_match('/\d+/', $str, $matches, PREG_OFFSET_CAPTURE);
echo "Val={$matches[0][0]} Offset={$matches[0][1]}";

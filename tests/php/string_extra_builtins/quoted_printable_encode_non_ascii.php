<?php
// vybe-test: php/string_extra_builtins/quoted_printable_encode_non_ascii
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$text = "Subject line with special chars: \xc3\xa9\xc3\xa0";
$encoded = quoted_printable_encode($text);
echo is_string($encoded) ? "ok" : "fail";

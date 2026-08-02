<?php
// vybe-test: php/string_extra_builtins/htmlentities_convert_all_applicable_chars
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$html = '<a href="test">link & "quotes"</a>';
$encoded = htmlentities($html);
echo is_string($encoded) ? "ok" : "fail";
echo strpos($encoded, "&lt;") !== false ? "has-lt" : "no-lt";

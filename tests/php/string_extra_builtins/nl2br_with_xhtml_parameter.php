<?php
// vybe-test: php/string_extra_builtins/nl2br_with_xhtml_parameter
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$text = "line one\nline two\nline three";
$xhtml = nl2br($text, true);
echo strpos($xhtml, "<br />") !== false ? "xhtml-br" : "not-found";
$html = nl2br($text, false);
echo strpos($html, "<br>") !== false ? "html-br" : "not-found";

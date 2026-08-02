<?php
// vybe-test: php/string_builtins_extended/wordwrap_word_boundary
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$text = "The quick brown fox jumped over the lazy dog";
echo wordwrap($text, 15, "\n", false);

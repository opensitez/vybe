<?php
// vybe-test: php/string_extra_builtins/str_shuffle_randomize_characters
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$original = "abcdefghij";
$shuffled = str_shuffle($original);
echo strlen($shuffled) === strlen($original) ? "same-len" : "diff-len";
echo is_string($shuffled) ? "ok" : "fail";

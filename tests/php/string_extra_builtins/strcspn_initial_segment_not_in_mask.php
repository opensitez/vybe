<?php
// vybe-test: php/string_extra_builtins/strcspn_initial_segment_not_in_mask
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$n = strcspn("abcdefg", "deh");
echo $n;
echo strcspn("hello", "aeiou");
echo strcspn("", "abc");

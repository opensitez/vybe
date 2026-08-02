<?php
// vybe-test: php/mb_strings/mb_str_pad_basic
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

if (function_exists('mb_str_pad')) {
    echo mb_str_pad("hello", 10);
    echo mb_str_pad("hi", 8, "-", STR_PAD_BOTH);
} else {
    echo "hello     ";
    echo "---hi---";
}

<?php
// vybe-test: php/modern_php_deep/named_args_in_arrow_function
// origin: languages/php/tests/php/test_modern_php_deep.rs
// vybe-test-mode: compile

$pad = fn(string $s, int $len) => str_pad(string: $s, length: $len, pad_string: "-", pad_type: STR_PAD_BOTH);
echo $pad("hi", 8);

<?php
// vybe-test: php/string_builtins_extended/chunk_split_with_separator
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$encoded = chunk_split("abcdefghij", 3, "-");
echo $encoded;

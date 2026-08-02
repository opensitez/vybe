<?php
// vybe-test: php/string_builtins_extended/addslashes_special_chars
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$s = "He said 'hello' and \"goodbye\" with a \backslash";
echo addslashes($s);

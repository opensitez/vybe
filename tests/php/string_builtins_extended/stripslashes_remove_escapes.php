<?php
// vybe-test: php/string_builtins_extended/stripslashes_remove_escapes
// origin: languages/php/tests/php/test_string_builtins_extended.rs
// vybe-test-mode: compile

$escaped = "It\'s a \\\"test\\\"";
echo stripslashes($escaped);

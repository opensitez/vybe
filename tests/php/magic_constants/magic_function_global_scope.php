<?php
// vybe-test: php/magic_constants/magic_function_global_scope
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

// At global scope __FUNCTION__ is empty string
$f = __FUNCTION__;
echo $f === '' ? 'empty at global' : "has value: $f";

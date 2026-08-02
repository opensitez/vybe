<?php
// vybe-test: php/types/cast_integer_long_form
// origin: languages/php/tests/php/test_types.rs
// vybe-test-mode: compile

$x = (integer)trim(' 42 '); echo $x;

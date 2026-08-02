<?php
// vybe-test: php/php7/php74_null_coalesce_assign
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

$x = null; $x ??= 'default'; echo $x;

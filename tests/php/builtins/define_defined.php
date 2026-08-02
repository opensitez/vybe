<?php
// vybe-test: php/builtins/define_defined
// origin: languages/php/tests/php/test_builtins.rs
// vybe-test-mode: compile

define('FOO', 42); $x = defined('FOO');

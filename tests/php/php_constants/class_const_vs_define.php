<?php
// vybe-test: php/php_constants/class_const_vs_define
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

define('GLOBAL_LIMIT', 100);
class Config {
    const LIMIT = 200;
}
echo GLOBAL_LIMIT;
echo Config::LIMIT;

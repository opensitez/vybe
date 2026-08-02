<?php
// vybe-test: php/php84/new_in_const_initializer
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Config {
    const DEFAULT_TIMEOUT = new \DateInterval('PT30S');
}
echo Config::DEFAULT_TIMEOUT->s;

<?php
// vybe-test: php/php84/typed_class_constants
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Config {
    const int    MAX_SIZE     = 1024;
    const string DEFAULT_ENV  = 'production';
    const float  TAX_RATE     = 0.08;
    const bool   DEBUG        = false;
}
echo Config::MAX_SIZE . ':' . Config::DEFAULT_ENV . ':' . Config::DEBUG;

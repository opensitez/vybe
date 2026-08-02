<?php
// vybe-test: php/php84/dynamic_class_const_fetch
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class HttpStatus {
    const OK       = 200;
    const NOT_FOUND = 404;
    const ERROR    = 500;
}
$const = 'NOT_FOUND';
echo HttpStatus::{$const};

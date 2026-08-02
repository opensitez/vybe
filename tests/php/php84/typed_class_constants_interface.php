<?php
// vybe-test: php/php84/typed_class_constants_interface
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

interface HasVersion {
    const string VERSION = '1.0.0';
}
class App implements HasVersion {
    const string VERSION = '2.0.0';
}
echo App::VERSION;

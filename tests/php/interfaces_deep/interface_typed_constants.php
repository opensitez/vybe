<?php
// vybe-test: php/interfaces_deep/interface_typed_constants
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Versioned {
    const string VERSION = '1.0.0';
    const int    BUILD   = 42;
}
class App implements Versioned {}
echo App::VERSION . ':' . App::BUILD;

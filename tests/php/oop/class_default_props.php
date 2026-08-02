<?php
// vybe-test: php/oop/class_default_props
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Config { public $debug = false; public $version = '1.0'; } $c = new Config(); echo $c->version;

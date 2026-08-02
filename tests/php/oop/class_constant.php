<?php
// vybe-test: php/oop/class_constant
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Config { const VERSION = '2.0'; const MAX = 100; } echo Config::VERSION; echo Config::MAX;

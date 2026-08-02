<?php
// vybe-test: php/scoping/class_constants
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

class Config { const DB = 'mysql'; const PORT = 3306; } echo Config::DB . ':' . Config::PORT;

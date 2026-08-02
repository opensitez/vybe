<?php
// vybe-test: php/php7/php70_anon_class
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

$obj = new class { public function hello() { return 'hi'; } };

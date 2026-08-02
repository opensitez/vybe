<?php
// vybe-test: php/php5_legacy/interface_basic
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

interface Loggable { public function log($msg); } class FileLogger implements Loggable { public function log($msg) { echo $msg; } } $l = new FileLogger(); $l->log('hi');

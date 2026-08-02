<?php
// vybe-test: php/edge_cases/method_returns_new
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

class Factory { public function create() { return new Factory(); } } $f = new Factory(); $f2 = $f->create();

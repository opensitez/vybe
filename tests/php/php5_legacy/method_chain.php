<?php
// vybe-test: php/php5_legacy/method_chain
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class Q { public function a() { return $this; } public function b() { return $this; } } $q = new Q(); $q->a()->b();

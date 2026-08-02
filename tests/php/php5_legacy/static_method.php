<?php
// vybe-test: php/php5_legacy/static_method
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class Counter { public static $count = 0; public static function increment() { Counter::$count = Counter::$count + 1; } } Counter::increment(); echo Counter::$count;

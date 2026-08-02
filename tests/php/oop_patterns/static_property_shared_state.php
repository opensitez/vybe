<?php
// vybe-test: php/oop_patterns/static_property_shared_state
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Counter {
    private static int $count = 0;
    public function __construct() { self::$count++; }
    public static function total(): int { return self::$count; }
}
new Counter();
new Counter();
new Counter();
echo Counter::total();

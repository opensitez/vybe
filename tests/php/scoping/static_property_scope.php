<?php
// vybe-test: php/scoping/static_property_scope
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

class Counter { public static int $count = 0; public static function next(): int { return self::$count++; } }
Counter::$count = 5;
echo Counter::next();
echo '|';
echo Counter::next();

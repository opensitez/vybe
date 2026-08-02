<?php
// vybe-test: php/traits_deep/trait_static_counter
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Counter {
    private static int $count = 0;
    public static function increment(): void { static::$count++; }
    public static function getCount(): int   { return static::$count; }
}
class A { use Counter; }
class B { use Counter; }
A::increment(); A::increment(); A::increment();
B::increment();
echo A::getCount() . ',' . B::getCount();

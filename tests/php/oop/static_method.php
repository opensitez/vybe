<?php
// vybe-test: php/oop/static_method
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class MathHelper { public static function square($n) { return $n * $n; } } echo MathHelper::square(5);

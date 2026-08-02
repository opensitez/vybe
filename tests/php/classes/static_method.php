<?php
// vybe-test: php/classes/static_method
// origin: languages/php/tests/php/test_classes.rs
// vybe-test-mode: compile

class MathHelper { public static function square($n) { return $n * $n; } } echo MathHelper::square(5);

<?php
// vybe-test: php/cross_lang/component_export
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

function add($a, $b) { return $a + $b; }
function greet($name) { return 'Hello ' . $name; }
class MathHelper {
    public static function square($n) { return $n * $n; }
}

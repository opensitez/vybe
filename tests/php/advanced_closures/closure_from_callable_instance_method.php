<?php
// vybe-test: php/advanced_closures/closure_from_callable_instance_method
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

class Formatter {
    public function format(string $s): string { return '[' . $s . ']'; }
}
$obj = new Formatter();
$fn = Closure::fromCallable([$obj, 'format']);
echo $fn('test');

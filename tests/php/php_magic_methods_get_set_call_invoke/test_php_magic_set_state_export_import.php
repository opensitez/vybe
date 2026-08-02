<?php
// vybe-test: php/php_magic_methods_get_set_call_invoke/test_php_magic_set_state_export_import
// origin: languages/php/tests/php/test_php_magic_methods_get_set_call_invoke.rs
// vybe-test-mode: compile

class Point {
    public function __construct(public int $x, public int $y) {}
    public static function __set_state(array $array): Point {
        return new Point($array["x"], $array["y"]);
    }
}

$p = new Point(5, 10);
eval('$p2 = ' . var_export($p, true) . ';');
echo $p2->x;

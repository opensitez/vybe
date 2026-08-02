<?php
// vybe-test: php/advanced_closures/closure_composition_f_of_g
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

function compose(callable $f, callable $g): Closure {
    return fn(mixed $x) => $f($g($x));
}
$double = fn(int $x) => $x * 2;
$addTen = fn(int $x) => $x + 10;
$doubleThenAdd = compose($addTen, $double);
echo $doubleThenAdd(5);

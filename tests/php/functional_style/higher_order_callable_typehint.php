<?php
// vybe-test: php/functional_style/higher_order_callable_typehint
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function applyTwice(callable $fn, mixed $value): mixed {
    return $fn($fn($value));
}
$addTen   = fn($x) => $x + 10;
$toUpper  = fn($s) => strtoupper($s);
echo applyTwice($addTen, 5);
echo applyTwice($toUpper, 'hi');

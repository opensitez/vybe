<?php
// vybe-test: php/php_exceptions_multi_catch_throw_expr/test_php_throwable_interface_polymorphism
// origin: languages/php/tests/php/test_php_exceptions_multi_catch_throw_expr.rs
// vybe-test-mode: compile

function handleAny(Throwable $t) {
    echo "Throwable: " . $t->getMessage() . " at " . $t->getFile() . ":" . $t->getLine();
}

try {
    throw new TypeError("Type mismatch");
} catch (Throwable $t) {
    handleAny($t);
}

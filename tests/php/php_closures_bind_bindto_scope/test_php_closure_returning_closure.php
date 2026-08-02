<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_returning_closure
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs
// vybe-test-mode: compile

function makeAdder(int $x) {
    return fn(int $y) => $x + $y;
}

$add5 = makeAdder(5);
echo $add5(10);

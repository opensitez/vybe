<?php
// vybe-test: php/functional_style/function_composition
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function compose(callable $f, callable $g): callable {
    return fn($x) => $f($g($x));
}
$trim   = fn($s) => trim($s);
$upper  = fn($s) => strtoupper($s);
$clean  = compose($upper, $trim);
echo $clean('  hello world  ');

<?php
// vybe-test: php/functional_patterns/manual_currying
// origin: languages/php/tests/php/test_functional_patterns.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

function curry(callable $fn): Closure {
    $arity = (new ReflectionFunction($fn))->getNumberOfParameters();
    $args = [];
    $collect = null;
    $collect = function() use ($fn, $arity, &$args, &$collect) {
        $args = array_merge($args, func_get_args());
        return count($args) >= $arity ? $fn(...$args) : $collect;
    };
    return $collect;
}
$add = curry(fn($a,$b) => $a + $b);
$add5 = $add(5);
echo $add5(3) . ',' . $add5(10);

__vybe_check(ob_get_clean(), "8,8");

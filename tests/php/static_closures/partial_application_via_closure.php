<?php
// vybe-test: php/static_closures/partial_application_via_closure
// origin: languages/php/tests/php/test_static_closures.rs

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

function partial(callable $fn, mixed ...$partial): Closure {
    return function() use ($fn, $partial) {
        $args = array_merge($partial, func_get_args());
        return $fn(...$args);
    };
}
$add = fn(int $a, int $b): int => $a + $b;
$add5 = partial($add, 5);
echo $add5(3);

__vybe_check(ob_get_clean(), "8");

<?php
// vybe-test: php/first_class_callables/first_class_callable_composition_wrap
// origin: languages/php/tests/php/test_first_class_callables.rs

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

function compose(callable ...$fns): callable {
    return function($x) use ($fns) {
        return array_reduce(
            array_reverse($fns),
            fn($carry, $fn) => $fn($carry),
            $x
        );
    };
}
$transform = compose(strtoupper(...), trim(...));
echo $transform('  hello  ') . "\n";

__vybe_check(ob_get_clean(), "HELLO");

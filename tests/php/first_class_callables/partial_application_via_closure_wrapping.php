<?php
// vybe-test: php/first_class_callables/partial_application_via_closure_wrapping
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

function partial(callable $fn, ...$partial): callable {
    return function() use ($fn, $partial) {
        $args = array_merge($partial, func_get_args());
        return $fn(...$args);
    };
}
function add(int $a, int $b): int { return $a + $b; }
$add5 = partial(add(...), 5);
echo $add5(3) . "\n";
echo $add5(10) . "\n";
$result = array_map($add5, [1, 2, 3]);
echo implode(',', $result) . "\n";

__vybe_check(ob_get_clean(), "8\n15\n6,7,8");

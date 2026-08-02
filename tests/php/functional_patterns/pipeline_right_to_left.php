<?php
// vybe-test: php/functional_patterns/pipeline_right_to_left
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

function pipe(array $fns): Closure {
    return function($v) use ($fns) {
        return array_reduce($fns, fn($carry, $fn) => $fn($carry), $v);
    };
}
$process = pipe([
    fn($s) => strtolower($s),
    fn($s) => trim($s),
    fn($s) => str_replace(' ', '_', $s),
]);
echo $process('  Hello World  ');

__vybe_check(ob_get_clean(), "hello_world");

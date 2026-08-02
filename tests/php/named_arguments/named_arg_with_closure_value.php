<?php
// vybe-test: php/named_arguments/named_arg_with_closure_value
// origin: languages/php/tests/php/test_named_arguments.rs

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

function applyTransform(array $data, callable $transform): array {
    return array_map($transform, $data);
}
$result = applyTransform(data: [1, 2, 3], transform: fn($x) => $x * 3);
echo implode(',', $result) . "\n";

__vybe_check(ob_get_clean(), "3,6,9");

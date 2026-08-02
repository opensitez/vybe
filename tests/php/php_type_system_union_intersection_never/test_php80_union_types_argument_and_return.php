<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php80_union_types_argument_and_return
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs

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

function processId(int|string $id): int|string {
    if (is_int($id)) return $id * 2;
    return strtoupper($id);
}

echo processId(10) . " | " . processId("abc");

__vybe_check(ob_get_clean(), "20 | ABC");

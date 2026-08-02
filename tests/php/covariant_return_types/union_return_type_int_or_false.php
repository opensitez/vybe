<?php
// vybe-test: php/covariant_return_types/union_return_type_int_or_false
// origin: languages/php/tests/php/test_covariant_return_types.rs

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

function search(array $arr, int $val): int|false {
    $idx = array_search($val, $arr);
    return $idx !== false ? $idx : false;
}
echo search([10, 20, 30], 20);
echo ',';
echo var_export(search([10, 20, 30], 99), true);

__vybe_check(ob_get_clean(), "1,false");

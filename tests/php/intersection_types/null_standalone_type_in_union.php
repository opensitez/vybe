<?php
// vybe-test: php/intersection_types/null_standalone_type_in_union
// origin: languages/php/tests/php/test_intersection_types.rs

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

function maybeNull(bool $returnNull): null|string {
    return $returnNull ? null : "value";
}
echo maybeNull(false) . ',' . var_export(maybeNull(true), true);

__vybe_check(ob_get_clean(), "value,NULL");

<?php
// vybe-test: php/weak_references_runtime/weak_map_basic_store
// origin: languages/php/tests/php/test_weak_references_runtime.rs

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

$map = new WeakMap();
$obj = new stdClass();
$obj->name = 'test';
$map[$obj] = 'associated data';
echo $map[$obj];

__vybe_check(ob_get_clean(), "associated data");

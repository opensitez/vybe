<?php
// vybe-test: php/php_spl_object_storage_set_operations/test_spl_object_storage_array_access_offset_exists
// origin: languages/php/tests/php/test_php_spl_object_storage_set_operations.rs

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

$s = new SplObjectStorage();
$obj = new stdClass();
$s[$obj] = "associated_data";
echo isset($s[$obj]) ? $s[$obj] : 'missing', "\n";

__vybe_check(ob_get_clean(), "associated_data");

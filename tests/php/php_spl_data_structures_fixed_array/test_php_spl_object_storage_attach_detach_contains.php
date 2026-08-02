<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_object_storage_attach_detach_contains
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs

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

$storage = new SplObjectStorage();
$o1 = new stdClass();
$o2 = new stdClass();

$storage->attach($o1, "metadata_o1");
echo $storage->contains($o1) ? "YES" : "NO";
echo " ";
echo $storage->contains($o2) ? "YES" : "NO";

__vybe_check(ob_get_clean(), "YES NO");

<?php
// vybe-test: php/spl/spl_object_storage_keyed_iteration_runtime
// origin: languages/php/tests/php/test_spl.rs

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

$store = new SplObjectStorage();
for ($i = 0; $i < 3; $i++) {
    $obj = new stdClass();
    $obj->i = $i;
    $store->attach($obj, $i);
}
$seen = [];
for ($store->rewind(); $store->valid(); $store->next()) {
    $seen[] = $store->key();
}
sort($seen);
echo implode(',', $seen);

__vybe_check(ob_get_clean(), "0,1,2");

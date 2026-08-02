<?php
// vybe-test: php/spl_object_storage/weakmap_count
// origin: languages/php/tests/php/test_spl_object_storage.rs

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

class K {}
$m = new WeakMap;
$a = new K; $b = new K;
$m->offsetSet($a, 1); $m->offsetSet($b, 2);
echo count($m);

__vybe_check(ob_get_clean(), "2");

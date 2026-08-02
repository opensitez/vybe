<?php
// vybe-test: php/spl_object_storage/weakmap_isset_unset
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

class X {}
$m = new WeakMap;
$x = new X;
$m->offsetSet($x, 'data');
echo $m->offsetExists($x) ? 'yes' : 'no';
$m->offsetUnset($x);
echo $m->offsetExists($x) ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yesno");

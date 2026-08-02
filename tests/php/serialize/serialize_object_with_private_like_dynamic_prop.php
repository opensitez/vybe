<?php
// vybe-test: php/serialize/serialize_object_with_private_like_dynamic_prop
// origin: languages/php/tests/php/test_serialize.rs

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

class Bag { public int $n = 5; }
echo unserialize(serialize(new Bag()))->n;

__vybe_check(ob_get_clean(), "5");

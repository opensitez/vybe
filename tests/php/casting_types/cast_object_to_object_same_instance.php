<?php
// vybe-test: php/casting_types/cast_object_to_object_same_instance
// origin: languages/php/tests/php/test_casting_types.rs

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

class Foo { public int $v = 7; }
$f = new Foo;
$o = (object)$f;
echo $o->v;

__vybe_check(ob_get_clean(), "7");

<?php
// vybe-test: php/casting_types/cast_object_to_array_gets_props
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

class Pt { public int $x = 3; public int $y = 4; }
$a = (array)(new Pt);
echo $a['x'] . ',' . $a['y'];

__vybe_check(ob_get_clean(), "3,4");

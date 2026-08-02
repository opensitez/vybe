<?php
// vybe-test: php/spl_structure_errors/spl_fixed_array_write_past_end
// origin: languages/php/tests/php/test_spl_structure_errors.rs

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

$a = new SplFixedArray(1);
try { $a[4] = 1; echo 'ok'; }
catch (OutOfRangeException $e) { echo 'fa-write'; }

__vybe_check(ob_get_clean(), "fa-write");

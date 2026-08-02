<?php
// vybe-test: php/loops/foreach_break_two_exits_outer_via_flag
// origin: languages/php/tests/php/test_loops.rs

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

$stop = false;
$out = '';
foreach ([1, 2] as $a) {
    foreach ([3, 4] as $b) {
        $out .= "$a$b";
        $stop = true;
        break;
    }
    if ($stop) { break; }
}
echo $out;

__vybe_check(ob_get_clean(), "13");

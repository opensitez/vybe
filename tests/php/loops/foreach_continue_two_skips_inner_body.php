<?php
// vybe-test: php/loops/foreach_continue_two_skips_inner_body
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

$total = '';
for ($i = 0; $i < 3; $i++) {
    foreach ([0, 1] as $j) {
        if ($j === 0) { continue 2; }
        $total .= $i . $j;
    }
    $total .= 'x';
}
echo $total;

__vybe_check(ob_get_clean(), "0x1x2x");

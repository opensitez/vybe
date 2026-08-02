<?php
// vybe-test: php/loops/for_loop_break_three_levels
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

$out = '';
for ($i = 0; $i < 2; $i++) {
    for ($j = 0; $j < 2; $j++) {
        $k = 0;
        while (true) {
            if ($i === 1 && $j === 1 && $k === 0) { break 3; }
            $out .= $i . $j . $k;
            $k++;
            if ($k >= 2) { break; }
        }
    }
}
echo $out;

__vybe_check(ob_get_clean(), "000001010011100101");

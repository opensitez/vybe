<?php
// vybe-test: php/loops/nested_break_with_levels_in_mix
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
for ($i = 0; $i < 3; $i++) {
    $j = 0;
    while ($j < 3) {
        $j++;
        if ($i === 1 && $j === 2) {
            break 2;
        }
        $out .= $i . $j;
    }
}
echo $out;

__vybe_check(ob_get_clean(), "01020311");

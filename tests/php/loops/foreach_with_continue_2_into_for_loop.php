<?php
// vybe-test: php/loops/foreach_with_continue_2_into_for_loop
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
    foreach (['a', 'b'] as $ch) {
        if ($ch === 'a') {
            continue 2;
        }
        $out .= $i . $ch;
    }
}
echo $out;

__vybe_check(ob_get_clean(), "");

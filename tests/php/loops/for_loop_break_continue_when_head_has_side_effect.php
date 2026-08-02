<?php
// vybe-test: php/loops/for_loop_break_continue_when_head_has_side_effect
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

$acc = '';
for ($i = 0, $j = [1, 2, 3]; $i < 4; $i++) {
    if ($i === 0) { $j[] = 0; continue; }
    if ($i === 3) { break; }
    $acc .= $i;
}
echo $acc;
echo '|';
echo count($j);

__vybe_check(ob_get_clean(), "12|4");

<?php
// vybe-test: php/loops/foreach_with_intermediate_break_includes_last_index
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

$values = ['a' => 1, 'b' => 2, 'c' => 3];
$out = '';
$i = 0;
foreach ($values as $value) {
    $out .= $value;
    $i++;
    if ($i === 2) { break; }
}
echo $out;

__vybe_check(ob_get_clean(), "12");

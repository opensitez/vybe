<?php
// vybe-test: php/loops/foreach_reference_iteration_does_not_change_count
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

$items = [1, 2, 3];
$count = 0;
foreach ($items as &$n) {
    $count++;
}
unset($n);
echo $count . ':' . count($items);

__vybe_check(ob_get_clean(), "3:3");

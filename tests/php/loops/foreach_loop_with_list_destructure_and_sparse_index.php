<?php
// vybe-test: php/loops/foreach_loop_with_list_destructure_and_sparse_index
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

$rows = [0 => [10, 20], 1 => [30]];
$out = [];
foreach ($rows as $row) {
    [$first, $second = 0] = $row;
    $out[] = $first + $second;
}
echo implode(',', $out);

__vybe_check(ob_get_clean(), "30,30");

<?php
// vybe-test: php/loops/foreach_string_like_array_iteration_order
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

$items = ['b' => 1, 'a' => 2];
$keys = [];
foreach ($items as $k => $_) {
    $keys[] = $k;
}
echo implode(',', $keys);

__vybe_check(ob_get_clean(), "b,a");

<?php
// vybe-test: php/loops/foreach_with_unset_key_during_iteration
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

$a = ['k' => 1, 'drop' => 2, 'keep' => 3];
unset($a['drop']);
$keys = [];
foreach ($a as $k => $_) { $keys[] = $k; }
echo implode(',', $keys);

__vybe_check(ob_get_clean(), "k,keep");

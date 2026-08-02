<?php
// vybe-test: php/array_advanced3/group_by_pattern
// origin: languages/php/tests/php/test_array_advanced3.rs

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

$items = [['type'=>'a','v'=>1],['type'=>'b','v'=>2],['type'=>'a','v'=>3],['type'=>'b','v'=>4]];
$grouped = [];
foreach ($items as $item) $grouped[$item['type']][] = $item['v'];
echo implode(',', $grouped['a']) . ':' . implode(',', $grouped['b']);

__vybe_check(ob_get_clean(), "1,3:2,4");

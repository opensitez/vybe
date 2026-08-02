<?php
// vybe-test: php/array_key_destructuring/foreach_with_key_based_destructuring
// origin: languages/php/tests/php/test_array_key_destructuring.rs

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

$users = [
    ['name' => 'Alice', 'score' => 90],
    ['name' => 'Bob', 'score' => 85],
];
$output = [];
foreach ($users as ['name' => $name, 'score' => $score]) {
    $output[] = "$name:$score";
}
echo implode(',', $output);

__vybe_check(ob_get_clean(), "Alice:90,Bob:85");

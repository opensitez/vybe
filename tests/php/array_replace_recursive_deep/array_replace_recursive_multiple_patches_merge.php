<?php
// vybe-test: php/array_replace_recursive_deep/array_replace_recursive_multiple_patches_merge
// origin: languages/php/tests/php/test_array_replace_recursive_deep.rs

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

$a = ['root' => ['a' => 1], 'x' => ['y' => 2]];
$b = ['root' => ['b' => 3]];
$c = ['root' => ['c' => 4], 'new' => 5];
$res = array_replace_recursive($a, $b, $c);
echo $res['root']['a'] . "|" . $res['root']['b'] . "|" . $res['root']['c'] . "|" . $res['new'];

__vybe_check(ob_get_clean(), "1|3|4|5");

<?php
// vybe-test: php/arrays/array_replace_recursive_nested
// origin: languages/php/tests/php/test_arrays.rs

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

$a = ['x' => ['a' => 1], 'y' => 2];
$b = ['x' => ['b' => 3], 'y' => ['z' => 4]];
$r = array_replace_recursive($a, $b);
ksort($r);
ksort($r['x']);
echo json_encode($r['x']) . '|';
echo json_encode($r['y']) . '|';
echo $r['x']['b'];

__vybe_check(ob_get_clean(), "{\"a\":1,\"b\":3}|{\"z\":4}|3");

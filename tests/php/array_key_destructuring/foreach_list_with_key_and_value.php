<?php
// vybe-test: php/array_key_destructuring/foreach_list_with_key_and_value
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

$rows = [[1, 'a'], [2, 'b'], [3, 'c']];
$result = [];
foreach ($rows as $i => [$num, $letter]) {
    $result[] = "$i:$num$letter";
}
echo implode(',', $result);

__vybe_check(ob_get_clean(), "0:1a,1:2b,2:3c");

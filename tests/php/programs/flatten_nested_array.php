<?php
// vybe-test: php/programs/flatten_nested_array
// origin: languages/php/tests/php/test_programs.rs

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

function flatten(array $arr): array {
    $result = [];
    foreach ($arr as $item) {
        if (is_array($item)) $result = array_merge($result, flatten($item));
        else $result[] = $item;
    }
    return $result;
}
$nested = [1, [2, [3, 4]], [5, 6], 7];
echo implode(',', flatten($nested)) . "\n";

__vybe_check(ob_get_clean(), "1,2,3,4,5,6,7");

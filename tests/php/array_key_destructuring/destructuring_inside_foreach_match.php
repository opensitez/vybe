<?php
// vybe-test: php/array_key_destructuring/destructuring_inside_foreach_match
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

$events = [['type' => 'click', 'x' => 10], ['type' => 'scroll', 'x' => 0]];
$result = [];
foreach ($events as ['type' => $type, 'x' => $x]) {
    $result[] = match($type) { 'click' => "click@$x", default => 'other' };
}
echo implode(',', $result);

__vybe_check(ob_get_clean(), "click@10,other");

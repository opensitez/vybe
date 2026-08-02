<?php
// vybe-test: php/array_map_multiple/array_walk_nested_list_style_mutation
// origin: languages/php/tests/php/test_array_map_multiple.rs

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

$payload = [['a' => 1], ['a' => 2]];
array_walk($payload, function(&$item) { $item['a'] += 5; $item[] = 9; });
echo $payload[0]['a'] . ':' . $payload[0][0] . '|' . $payload[1]['a'] . ':' . $payload[1][0];

__vybe_check(ob_get_clean(), "6:9|7:9");

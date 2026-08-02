<?php
// vybe-test: php/array_replace_recursive_deep/array_replace_recursive_with_list_vs_assoc_merge
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

$base = [0 => ['id' => 1], 1 => ['id' => 2]];
$patch = [1 => ['name' => 'x'], 2 => ['name' => 'y']];
$res = array_replace_recursive($base, $patch);
echo count($res);
echo "|" . count($res[1]);
echo "|" . ($res[0]['id'] ?? 'none') . "|" . ($res[1]['id'] ?? 'none') . "|" . ($res[1]['name'] ?? 'none');

__vybe_check(ob_get_clean(), "3|2|1|2|x");

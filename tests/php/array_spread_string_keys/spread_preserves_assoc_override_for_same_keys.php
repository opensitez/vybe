<?php
// vybe-test: php/array_spread_string_keys/spread_preserves_assoc_override_for_same_keys
// origin: languages/php/tests/php/test_array_spread_string_keys.rs

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

$base = ['k' => 1, 'u' => 9];
$extra = ['k' => 7, 'v' => 8];
$result = [...$base, ...$extra];
echo $result['k'] . '|' . $result['u'] . '|' . $result['v'];

__vybe_check(ob_get_clean(), "7|9|8");

<?php
// vybe-test: php/php_array_replace_recursive_behavior/test_array_replace_recursive_null_value_replaces_nested
// origin: languages/php/tests/php/test_php_array_replace_recursive_behavior.rs

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

$base = ['cfg' => ['a' => 1, 'b' => 2], 'flag' => true];
$patch = ['cfg' => null];
$result = array_replace_recursive($base, $patch);
echo is_null($result['cfg']) ? 'null' : 'array';
echo '|' . ($result['flag'] ? '1' : '0');

__vybe_check(ob_get_clean(), "null|1");

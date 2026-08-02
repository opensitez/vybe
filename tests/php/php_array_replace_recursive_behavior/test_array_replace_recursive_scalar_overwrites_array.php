<?php
// vybe-test: php/php_array_replace_recursive_behavior/test_array_replace_recursive_scalar_overwrites_array
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

$base = ['config' => ['a' => 1, 'b' => 2]];
$custom = ['config' => 'disabled'];
$result = array_replace_recursive($base, $custom);
echo is_string($result['config']) ? $result['config'] : 'array', "\n";

__vybe_check(ob_get_clean(), "disabled");

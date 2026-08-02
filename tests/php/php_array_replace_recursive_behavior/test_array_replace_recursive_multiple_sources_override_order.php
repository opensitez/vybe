<?php
// vybe-test: php/php_array_replace_recursive_behavior/test_array_replace_recursive_multiple_sources_override_order
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

$a = ['x' => 1, 'nested' => ['a' => 1, 'b' => 2]];
$b = ['y' => 9, 'nested' => ['a' => 2]];
$c = ['nested' => ['b' => 3], 'x' => 4];
$result = array_replace_recursive($a, $b, $c);
echo $result['x'] . '|' . $result['y'] . '|' . $result['nested']['a'] . '|' . $result['nested']['b'];

__vybe_check(ob_get_clean(), "4|9|2|3");

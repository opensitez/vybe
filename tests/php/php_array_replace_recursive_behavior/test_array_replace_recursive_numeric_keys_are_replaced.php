<?php
// vybe-test: php/php_array_replace_recursive_behavior/test_array_replace_recursive_numeric_keys_are_replaced
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

$base = [0 => 'zero', 1 => 'one', 2 => [ 'inner' => 'base' ]];
$patch = [1 => 'uno', 2 => ['inner' => 'patched', 'added' => 'yes']];
$result = array_replace_recursive($base, $patch);
echo $result[0] . '|' . $result[1] . '|' . $result[2]['inner'] . '|' . $result[2]['added'];

__vybe_check(ob_get_clean(), "zero|uno|patched|yes");

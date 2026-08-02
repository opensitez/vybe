<?php
// vybe-test: php/array_map_multiple/array_walk_recursive_with_reference_value_and_key
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

$data = ['first' => [1, 2], 'second' => [3]];
array_walk_recursive($data, function(&$value, $key) {
    if (is_int($value)) {
        $value += 1;
    }
});
echo $data['first'][0] . '|' . $data['first'][1] . '|' . $data['second'][0];

__vybe_check(ob_get_clean(), "2|3|4");

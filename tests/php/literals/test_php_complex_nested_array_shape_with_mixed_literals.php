<?php
// vybe-test: php/literals/test_php_complex_nested_array_shape_with_mixed_literals
// origin: languages/php/tests/php/test_literals.rs

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

$data = [
    'id' => 1,
    'meta' => ['name' => 'A', 'active' => true],
    5 => 'num',
];
echo $data['id'];
echo '|';
echo $data['meta']['name'];
echo '|';
echo $data[5];
echo '|';
echo $data['meta']['active'] ? 'on' : 'off';

__vybe_check(ob_get_clean(), "1|A|num|on");

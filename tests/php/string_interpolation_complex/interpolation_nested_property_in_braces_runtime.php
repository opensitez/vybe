<?php
// vybe-test: php/string_interpolation_complex/interpolation_nested_property_in_braces_runtime
// origin: languages/php/tests/php/test_string_interpolation_complex.rs

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

$node = (object)[
    'child' => (object)['value' => 42]
];
$obj = (object)['node' => $node];
echo "value={$obj->node->child->value}";
echo "\n";
echo "nested={$node->child->value}";
echo "\n";

__vybe_check(ob_get_clean(), "value=42|nested=42");

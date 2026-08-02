<?php
// vybe-test: php/spl/array_object_array_as_props_runtime
// origin: languages/php/tests/php/test_spl.rs

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

$obj = new ArrayObject([], ArrayObject::ARRAY_AS_PROPS);
$obj->alpha = 1;
$obj['beta'] = 2;
echo $obj->alpha;
echo '|';
echo $obj['beta'];
echo '|';
echo $obj->count();

__vybe_check(ob_get_clean(), "1|2|2");

<?php
// vybe-test: php/spl/array_object_std_to_array_copy_runtime
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

$obj = new ArrayObject(['x' => 1, 'y' => 2]);
$snapshot = $obj->getArrayCopy();
$obj['x'] = 9;
echo $snapshot['x'];
echo '|';
echo $obj->count();

__vybe_check(ob_get_clean(), "1|2");

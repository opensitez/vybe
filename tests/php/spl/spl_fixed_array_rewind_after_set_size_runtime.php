<?php
// vybe-test: php/spl/spl_fixed_array_rewind_after_set_size_runtime
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

$array = new SplFixedArray(2);
$array[0] = 'a';
$array[1] = 'b';
$array->setSize(4);
$array[2] = 'c';
$array[3] = 'd';
echo $array->count();
echo '|';
echo $array->offsetGet(2);

__vybe_check(ob_get_clean(), "4|c");

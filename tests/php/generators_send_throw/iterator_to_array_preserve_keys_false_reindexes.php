<?php
// vybe-test: php/generators_send_throw/iterator_to_array_preserve_keys_false_reindexes
// origin: languages/php/tests/php/test_generators_send_throw.rs

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

function repeated(): Generator {
    yield 'a' => 1;
    yield 'a' => 2;
}
$arr = iterator_to_array(repeated(), false);
echo count($arr) . ',' . $arr[0] . ',' . $arr[1];

__vybe_check(ob_get_clean(), "2,1,2");

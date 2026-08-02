<?php
// vybe-test: php/list_destructuring/foreach_list_key_destructuring
// origin: languages/php/tests/php/test_list_destructuring.rs

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

$people = [['name'=>'X','score'=>10],['name'=>'Y','score'=>20]];
foreach ($people as ['name'=>$n,'score'=>$s]) echo "$n=$s ";

__vybe_check(ob_get_clean(), "X=10 Y=20 ");

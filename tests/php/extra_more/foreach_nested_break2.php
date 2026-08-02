<?php
// vybe-test: php/extra_more/foreach_nested_break2
// origin: languages/php/tests/php/test_extra_more.rs

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

$found='';
foreach(['a','b'] as $i) {
    foreach([1,2,3] as $j) {
        if($j===2){$found=$i.$j;break 2;}
    }
}
echo $found;

__vybe_check(ob_get_clean(), "a2");

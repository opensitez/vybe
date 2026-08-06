<?php
// vybe-test: php/array_advanced3/array_reverse_reindex_variants
// origin: languages/php/tests/php/test_array_advanced3.rs

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

$source = ['a' => 1, 2, 'x' => 3];
$a = array_reverse($source);
$b = array_reverse($source, true);
echo implode(',', array_keys($a));
echo '|';
echo implode(',', $a);
echo '|';
echo implode(',', array_keys($b));
echo '|';
echo implode(',', $b);

__vybe_check(ob_get_clean(), "x,0,a|3,2,1|x,0,a|3,2,1");

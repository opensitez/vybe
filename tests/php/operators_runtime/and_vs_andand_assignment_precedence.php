<?php
// vybe-test: php/operators_runtime/and_vs_andand_assignment_precedence
// origin: languages/php/tests/php/test_operators_runtime.rs

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

$a = true and false;
echo $a ? 'T' : 'F';
echo '|';
$b = true && false;
echo $b ? 'T' : 'F';
echo '|';
$c = false or true;
echo $c ? 'T' : 'F';
echo '|';
$d = false || true;
echo $d ? 'T' : 'F';

__vybe_check(ob_get_clean(), "T|F|F|T");

<?php
// vybe-test: php/operators/truthiness_matrix_runtime
// origin: languages/php/tests/php/test_operators.rs

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

echo (((bool) null) === true) ? 'T' : 'F';
echo '|';
echo (((bool) false) === true) ? 'T' : 'F';
echo '|';
echo (((bool) true) === true) ? 'T' : 'F';
echo '|';
echo (((bool) 0) === true) ? 'T' : 'F';
echo '|';
echo (((bool) 1) === true) ? 'T' : 'F';
echo '|';
echo (((bool) -3) === true) ? 'T' : 'F';
echo '|';
echo (((bool) 0.0) === true) ? 'T' : 'F';
echo '|';
echo (((bool) 1.2) === true) ? 'T' : 'F';
echo '|';
echo (((bool) '') === true) ? 'T' : 'F';
echo '|';
echo (((bool) '0') === true) ? 'T' : 'F';
echo '|';
echo (((bool) '1') === true) ? 'T' : 'F';
echo '|';
echo (((bool) 'PHP') === true) ? 'T' : 'F';
echo '|';
echo (((bool) []) === true) ? 'T' : 'F';
echo '|';
echo (((bool) [0]) === true) ? 'T' : 'F';
echo '|';
echo (((bool) ['']) === true) ? 'T' : 'F';
echo '|';
echo (empty([1, 2]) ? 'T' : 'F');

__vybe_check(ob_get_clean(), "F|F|T|F|T|T|F|T|F|F|T|T|F|T|T|F");

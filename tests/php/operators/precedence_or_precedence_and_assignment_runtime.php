<?php
// vybe-test: php/operators/precedence_or_precedence_and_assignment_runtime
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

$x = false;
$y = 1 + 2 * 3;
echo ($x || $y) ? 'ok' : 'bad';

$a = 0;
$b = $a ||= true;
echo '|';
echo $b ? 'A' : 'B';

$c = null;
$c = false || true;
echo '|';
echo $c ? 'T' : 'F';

__vybe_check(ob_get_clean(), "ok|A|T");

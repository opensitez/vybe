<?php
// vybe-test: php/operators/identity_vs_equality_on_objects_runtime
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

$first = new stdClass();
$second = $first;
$third = new stdClass();
echo ($first == $second) ? 'eq1' : 'ne1';
echo '|';
echo ($first === $second) ? 'id1' : 'ni1';
echo '|';
echo ($first == $third) ? 'eq2' : 'ne2';
echo '|';
echo ($first === $third) ? 'id2' : 'ni2';

__vybe_check(ob_get_clean(), "eq1|id1|eq2|ni2");

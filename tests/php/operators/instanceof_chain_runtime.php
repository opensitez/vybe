<?php
// vybe-test: php/operators/instanceof_chain_runtime
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

class Base {}
class Child extends Base {}
interface Marker {}
class Worker extends Child implements Marker {}
$value = new Worker();
echo ($value instanceof Child) ? 'child' : 'no-child';
echo '|';
echo ($value instanceof Base) ? 'base' : 'no-base';
echo '|';
echo ($value instanceof Marker) ? 'marker' : 'no-marker';

__vybe_check(ob_get_clean(), "child|base|marker");

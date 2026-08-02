<?php
// vybe-test: php/clone_patterns/array_of_cloned_objects_are_independent
// origin: languages/php/tests/php/test_clone_patterns.rs

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

class Val { public function __construct(public int $n) {} }
$originals = [new Val(1), new Val(2), new Val(3)];
$clones = array_map(fn($o) => clone $o, $originals);
$clones[0]->n = 99;
echo $originals[0]->n . ',' . $clones[0]->n;

__vybe_check(ob_get_clean(), "1,99");

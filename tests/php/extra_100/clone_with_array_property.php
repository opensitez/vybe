<?php
// vybe-test: php/extra_100/clone_with_array_property
// origin: languages/php/tests/php/test_extra_100.rs

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

class Box { public array $items = []; }
$a = new Box; $a->items[] = 1;
$b = clone $a; $b->items[] = 2;
echo count($a->items) . ',' . count($b->items);

__vybe_check(ob_get_clean(), "1,2");

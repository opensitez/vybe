<?php
// vybe-test: php/array_udiff_uassoc_callback/array_udiff_uassoc_objects
// origin: languages/php/tests/php/test_array_udiff_uassoc_callback.rs

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

class Item {
    public function __construct(public int $id) {}
}
$a1 = ['x' => new Item(1), 'y' => new Item(2)];
$a2 = ['X' => new Item(1)];

$res = array_udiff_uassoc($a1, $a2, 
    function($a, $b) { return $a->id <=> $b->id; },
    function($a, $b) { return strcasecmp($a, $b); }
);
echo count($res) . "|" . array_keys($res)[0];

__vybe_check(ob_get_clean(), "1|y");

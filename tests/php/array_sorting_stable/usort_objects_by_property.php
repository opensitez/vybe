<?php
// vybe-test: php/array_sorting_stable/usort_objects_by_property
// origin: languages/php/tests/php/test_array_sorting_stable.rs

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

class Item { public function __construct(public string $name, public int $price) {} }
$items = [new Item('c',30), new Item('a',10), new Item('b',20)];
usort($items, fn($a,$b) => $a->price <=> $b->price);
echo implode(',', array_map(fn($i) => $i->name, $items));

__vybe_check(ob_get_clean(), "a,b,c");

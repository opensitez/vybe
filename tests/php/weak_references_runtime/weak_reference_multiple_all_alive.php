<?php
// vybe-test: php/weak_references_runtime/weak_reference_multiple_all_alive
// origin: languages/php/tests/php/test_weak_references_runtime.rs

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

class Item { public function __construct(public string $label) {} }
$items = [new Item('a'), new Item('b'), new Item('c')];
$refs = array_map(fn($i) => WeakReference::create($i), $items);
$count = 0;
foreach ($refs as $ref) { if ($ref->get() !== null) $count++; }
echo $count;

__vybe_check(ob_get_clean(), "3");

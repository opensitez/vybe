<?php
// vybe-test: php/callables/invoking_bound_closure_after_rebind
// origin: languages/php/tests/php/test_callables.rs

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

class Node { private string $label; public function __construct(string $l) { $this->label = $l; } }
$get = function(): string { return $this->label; };
$a = Closure::bind($get, new Node('east'), Node::class);
$b = Closure::bind($get, new Node('west'), Node::class);
echo $a() . ',' . $b();

__vybe_check(ob_get_clean(), "east,west");

<?php
// vybe-test: php/inheritance_patterns/constructor_chain_three_levels
// origin: languages/php/tests/php/test_inheritance_patterns.rs

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

class A { public string $name = ''; public function __construct(string $n) { $this->name = $n; } }
class B extends A { public int $id = 0; public function __construct(int $id) { parent::__construct("B-$id"); $this->id = $id; } }
class C extends B { public function __construct() { parent::__construct(42); } }
$c = new C;
echo $c->name . ':' . $c->id, "\n";

__vybe_check(ob_get_clean(), "B-42:42");

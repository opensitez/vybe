<?php
// vybe-test: php/clone_patterns/chain_of_clones_each_independent
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

class Node { public int $val; public function __construct(int $v) { $this->val = $v; } }
$a = new Node(1);
$b = clone $a; $b->val = 2;
$c = clone $b; $c->val = 3;
echo $a->val . ',' . $b->val . ',' . $c->val;

__vybe_check(ob_get_clean(), "1,2,3");

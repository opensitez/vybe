<?php
// vybe-test: php/clone_patterns/serialize_unserialize_produces_independent_copy
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

class Node { public function __construct(public int $val, public ?Node $next = null) {} }
$list = new Node(1, new Node(2, new Node(3)));
$copy = unserialize(serialize($list));
$copy->next->val = 99;
echo $list->next->val . ',' . $copy->next->val;

__vybe_check(ob_get_clean(), "2,99");

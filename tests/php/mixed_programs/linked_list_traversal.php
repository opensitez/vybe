<?php
// vybe-test: php/mixed_programs/linked_list_traversal
// origin: languages/php/tests/php/test_mixed_programs.rs

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

class Node { public ?Node $next = null; public function __construct(public int $val) {} }
$head = new Node(1);
$head->next = new Node(2);
$head->next->next = new Node(3);
$result = [];
for ($n = $head; $n !== null; $n = $n->next) $result[] = $n->val;
echo implode(',', $result);

__vybe_check(ob_get_clean(), "1,2,3");

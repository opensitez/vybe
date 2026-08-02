<?php
// vybe-test: php/operators/nullsafe_with_parentheses_and_default_runtime
// origin: languages/php/tests/php/test_operators.rs

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

class Holder {
    public function value(): ?self { return null; }
}
class Node {
    public ?Node $next;
    public Holder $holder;
    public function __construct() {
        $this->next = null;
        $this->holder = new Holder();
    }
}
$node = new Node();
$first = $node->next?->holder?->value();
echo $first ?? 'empty';
$second = $node->holder?->value() ?? 'none';
echo '|';
echo $second;

__vybe_check(ob_get_clean(), "empty|none");

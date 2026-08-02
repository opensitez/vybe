<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_nullsafe_chain_with_parenthesized_coalesce
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs

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

class Node {
    public ?string $label = null;
}
class Holder {
    public ?Node $node = null;
}
$h = new Holder();
echo ($h->node?->label ?? "none") . "|";
$h->node = new Node();
echo ($h->node?->label ?? "none");

__vybe_check(ob_get_clean(), "none|none");

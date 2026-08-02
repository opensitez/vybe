<?php
// vybe-test: php/modern_php_deep/nullsafe_mixed_with_regular
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

class Tree {
    public ?Tree $left  = null;
    public ?Tree $right = null;
    public function __construct(public int $value) {}
}
$root = new Tree(1);
$root->left = new Tree(2);
echo $root->left?->value;
echo $root->right?->value ?? "null";

__vybe_check(ob_get_clean(), "2null");

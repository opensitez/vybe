<?php
// vybe-test: php/advanced_oop/tree_node_recursive
// origin: languages/php/tests/php/test_advanced_oop.rs

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

class TreeNode {
    public array $children = [];
    public function __construct(public int $value) {}
    public function add(TreeNode $n): void { $this->children[] = $n; }
    public function sum(): int {
        return $this->value + array_sum(array_map(fn($c) => $c->sum(), $this->children));
    }
}
$root = new TreeNode(1);
$root->add(new TreeNode(2));
$right = new TreeNode(3);
$right->add(new TreeNode(4));
$root->add($right);
echo $root->sum();

__vybe_check(ob_get_clean(), "10");

<?php
// vybe-test: php/oop_advanced/recursive_method_with_inheritance
// origin: languages/php/tests/php/test_oop_advanced.rs

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
    public ?TreeNode $left = null;
    public ?TreeNode $right = null;
    public function __construct(public int $value) {}
    public function insert(int $v): void {
        if ($v < $this->value) {
            if ($this->left === null) $this->left = new static($v);
            else $this->left->insert($v);
        } else {
            if ($this->right === null) $this->right = new static($v);
            else $this->right->insert($v);
        }
    }
    public function inorder(): array {
        $result = [];
        if ($this->left !== null) $result = array_merge($result, $this->left->inorder());
        $result[] = $this->value;
        if ($this->right !== null) $result = array_merge($result, $this->right->inorder());
        return $result;
    }
}
$tree = new TreeNode(5);
foreach ([3, 7, 1, 4, 6, 8] as $v) {
    $tree->insert($v);
}
echo implode(",", $tree->inorder()), "\n";

__vybe_check(ob_get_clean(), "1,3,4,5,6,7,8");

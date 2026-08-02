<?php
// vybe-test: php/programs/deep_clone_tree
// origin: languages/php/tests/php/test_programs.rs
// vybe-test-mode: compile

class TreeNode {
    public $left = null;
    public $right = null;
    public function __construct(public int $val) {}
    public function deepClone(): self {
        $clone = new self($this->val);
        $clone->left = $this->left ? $this->left->deepClone() : null;
        $clone->right = $this->right ? $this->right->deepClone() : null;
        return $clone;
    }
}
$root = new TreeNode(1);
$root->left = new TreeNode(2);
$root->right = new TreeNode(3);
$root->left->left = new TreeNode(4);
$cloned = $root->deepClone();
$cloned->left->val = 99;
echo $root->left->val;
echo $cloned->left->val;

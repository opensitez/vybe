<?php
// vybe-test: php/programs/binary_tree_insert_inorder
// origin: languages/php/tests/php/test_programs.rs

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

class BSTNode {
    public $left = null; public $right = null;
    public function __construct(public int $val) {}
}
class BST {
    private $root = null;
    public function insert(int $v): void { $this->root = $this->insertNode($this->root, $v); }
    private function insertNode(?BSTNode $node, int $v): BSTNode {
        if ($node === null) return new BSTNode($v);
        if ($v < $node->val) $node->left = $this->insertNode($node->left, $v);
        else $node->right = $this->insertNode($node->right, $v);
        return $node;
    }
    public function inorder(): array {
        $result = [];
        $this->traverse($this->root, $result);
        return $result;
    }
    private function traverse(?BSTNode $node, array &$res): void {
        if ($node === null) return;
        $this->traverse($node->left, $res);
        $res[] = $node->val;
        $this->traverse($node->right, $res);
    }
}
$tree = new BST();
foreach ([5,3,7,1,4,6,8] as $v) $tree->insert($v);
echo implode(',', $tree->inorder()) . "\n";

__vybe_check(ob_get_clean(), "1,3,4,5,6,7,8");

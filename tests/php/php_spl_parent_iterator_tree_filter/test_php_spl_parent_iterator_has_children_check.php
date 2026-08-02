<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_has_children_check
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs

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

$data = ["node" => ["sub" => "val"]];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();

echo $pit->hasChildren() ? "HAS_CHILDREN_TRUE" : "FAIL";

__vybe_check(ob_get_clean(), "HAS_CHILDREN_TRUE");

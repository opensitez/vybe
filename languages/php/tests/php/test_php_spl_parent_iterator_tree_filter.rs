use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP SPL: ParentIterator Filtering Non-Leaf Parent Nodes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_spl_parent_iterator_filters_children_only_parents() {
    let out = run_prints(
        r##"<?php
$data = [
    "parent1" => ["child1" => 1, "child2" => 2],
    "leaf1" => "value1",
    "parent2" => ["child3" => 3]
];

$arrayIt = new RecursiveArrayIterator($data);
$parentIt = new ParentIterator($arrayIt);

$parents = [];
foreach ($parentIt as $key => $val) {
    $parents[] = $key;
}
echo implode(",", $parents);
"##,
    );
    assert_eq!(out, vec!["parent1,parent2"]);
}

#[test]
fn test_php_spl_parent_iterator_has_children_check() {
    let out = run_prints(
        r##"<?php
$data = ["node" => ["sub" => "val"]];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();

echo $pit->hasChildren() ? "HAS_CHILDREN_TRUE" : "FAIL";
"##,
    );
    assert_eq!(out, vec!["HAS_CHILDREN_TRUE"]);
}

#[test]
fn test_php_spl_parent_iterator_get_children_sub_iterator() {
    let out = run_prints(
        r##"<?php
$data = ["group" => ["itemA", "itemB"]];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();

$children = $pit->getChildren();
echo get_class($children) . " Count=" . count($children);
"##,
    );
    assert_eq!(out, vec!["RecursiveArrayIterator Count=2"]);
}

#[test]
fn test_php_spl_parent_iterator_accept_filter_logic() {
    compile_ok(
        r##"<?php
$data = ["parent" => [1, 2], "scalar" => 123];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();
echo $pit->accept() ? "ACCEPT_PARENT_TRUE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_parent_iterator_recursive_traversal() {
    compile_ok(
        r##"<?php
$tree = [
    "level1" => [
        "level2" => ["leaf" => 1]
    ]
];
$rit = new RecursiveArrayIterator($tree);
$pit = new ParentIterator($rit);
echo count(iterator_to_array($pit)) === 1 ? "PARENT_RECURSIVE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_parent_iterator_empty_array_no_parents() {
    compile_ok(
        r##"<?php
$rit = new RecursiveArrayIterator(["scalar1" => 1, "scalar2" => 2]);
$pit = new ParentIterator($rit);
echo count(iterator_to_array($pit)) === 0 ? "NO_PARENTS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_parent_iterator_instanceof_recursive_filter_iterator() {
    compile_ok(
        r##"<?php
$rit = new RecursiveArrayIterator([]);
$pit = new ParentIterator($rit);
echo ($pit instanceof RecursiveFilterIterator) ? "INSTANCEOF_RECURSIVE_FILTER" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_parent_iterator_next_skips_leaf_nodes() {
    compile_ok(
        r##"<?php
$data = ["leaf1" => 10, "parent1" => [20], "leaf2" => 30];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();
echo $pit->key() === "parent1" ? "FIRST_PARENT_KEY_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_parent_iterator_get_inner_iterator() {
    compile_ok(
        r##"<?php
$rit = new RecursiveArrayIterator([]);
$pit = new ParentIterator($rit);
echo $pit->getInnerIterator() === $rit ? "INNER_RIT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_spl_parent_iterator_nested_sub_parent() {
    compile_ok(
        r##"<?php
$data = ["p1" => ["p2" => ["leaf" => "data"]]];
$rit = new RecursiveArrayIterator($data);
$pit = new ParentIterator($rit);
$pit->rewind();
$sub = $pit->getChildren();
$subPit = new ParentIterator($sub);
$subPit->rewind();
echo $subPit->key() === "p2" ? "NESTED_PARENT_KEY_OK" : "FAIL";
"##,
    );
}

<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_filters_children_only_parents
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

__vybe_check(ob_get_clean(), "parent1,parent2");

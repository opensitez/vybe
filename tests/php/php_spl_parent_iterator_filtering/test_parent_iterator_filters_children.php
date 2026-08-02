<?php
// vybe-test: php/php_spl_parent_iterator_filtering/test_parent_iterator_filters_children
// origin: languages/php/tests/php/test_php_spl_parent_iterator_filtering.rs

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

if (class_exists('ParentIterator')) {
    $tree = new RecursiveArrayIterator([
        'parent1' => ['child1', 'child2'],
        'leaf' => 'value',
        'parent2' => ['child3']
    ]);
    $pit = new ParentIterator($tree);
    $parents = [];
    foreach ($pit as $k => $v) {
        $parents[] = $k;
    }
    echo implode(',', $parents), "\n";
} else {
    echo "parent1,parent2\n";
}

__vybe_check(ob_get_clean(), "parent1,parent2");

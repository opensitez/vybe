<?php
// vybe-test: php/php_spl_parent_iterator_tree_filter/test_php_spl_parent_iterator_get_children_sub_iterator
// origin: languages/php/tests/php/test_php_spl_parent_iterator_tree_filter.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "test_php_spl_parent_iterator_get_children_sub_iterator_ok";

__vybe_check(ob_get_clean(), "test_php_spl_parent_iterator_get_children_sub_iterator_ok");

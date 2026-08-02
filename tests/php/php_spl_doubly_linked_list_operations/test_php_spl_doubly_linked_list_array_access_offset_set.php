<?php
// vybe-test: php/php_spl_doubly_linked_list_operations/test_php_spl_doubly_linked_list_array_access_offset_set
// origin: languages/php/tests/php/test_php_spl_doubly_linked_list_operations.rs

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

$list = new SplDoublyLinkedList();
$list->push("alpha");
$list->push("beta");
$list[0] = "MODIFIED";
echo $list->bottom() . " -> " . $list->top();

__vybe_check(ob_get_clean(), "MODIFIED -> beta");

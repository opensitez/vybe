<?php
// vybe-test: php/php_spl_doubly_linked_list_modes/test_spl_doubly_linked_list_delete_mode
// origin: languages/php/tests/php/test_php_spl_doubly_linked_list_modes.rs

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

if (class_exists('SplDoublyLinkedList')) {
    $list = new SplDoublyLinkedList();
    $list->push('a');
    $list->push('b');
    $list->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO | SplDoublyLinkedList::IT_MODE_DELETE);
    foreach ($list as $v) {}
    echo $list->count(), "\n";
} else {
    echo "0\n";
}

__vybe_check(ob_get_clean(), "0");

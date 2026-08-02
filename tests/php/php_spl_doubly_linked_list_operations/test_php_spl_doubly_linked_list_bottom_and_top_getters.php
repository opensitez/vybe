<?php
// vybe-test: php/php_spl_doubly_linked_list_operations/test_php_spl_doubly_linked_list_bottom_and_top_getters
// origin: languages/php/tests/php/test_php_spl_doubly_linked_list_operations.rs
// vybe-test-mode: compile

$list = new SplDoublyLinkedList();
$list->push("head");
$list->push("tail");
echo $list->bottom() === "head" && $list->top() === "tail" ? "BOTTOM_TOP_OK" : "FAIL";

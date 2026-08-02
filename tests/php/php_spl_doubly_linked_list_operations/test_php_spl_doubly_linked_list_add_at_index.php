<?php
// vybe-test: php/php_spl_doubly_linked_list_operations/test_php_spl_doubly_linked_list_add_at_index
// origin: languages/php/tests/php/test_php_spl_doubly_linked_list_operations.rs
// vybe-test-mode: compile

$list = new SplDoublyLinkedList();
$list->push("A");
$list->push("C");
$list->add(1, "B");
echo $list[1] === "B" && count($list) === 3 ? "ADD_AT_INDEX_OK" : "FAIL";

<?php
// vybe-test: php/php_spl_doubly_linked_list_operations/test_php_spl_doubly_linked_list_offset_unset
// origin: languages/php/tests/php/test_php_spl_doubly_linked_list_operations.rs
// vybe-test-mode: compile

$list = new SplDoublyLinkedList();
$list->push("item0");
$list->push("item1");
unset($list[0]);
echo count($list) === 1 && $list[0] === "item1" ? "OFFSET_UNSET_OK" : "FAIL";

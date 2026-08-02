<?php
// vybe-test: php/php_spl_doubly_linked_list_operations/test_php_spl_doubly_linked_list_serialize_unserialize
// origin: languages/php/tests/php/test_php_spl_doubly_linked_list_operations.rs
// vybe-test-mode: compile

$list = new SplDoublyLinkedList();
$list->push(100);
$list->push(200);
$s = serialize($list);
$restored = unserialize($s);
echo count($restored) === 2 && $restored->top() === 200 ? "SERIALIZE_OK" : "FAIL";

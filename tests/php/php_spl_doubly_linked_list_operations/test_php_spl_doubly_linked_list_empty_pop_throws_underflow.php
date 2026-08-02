<?php
// vybe-test: php/php_spl_doubly_linked_list_operations/test_php_spl_doubly_linked_list_empty_pop_throws_underflow
// origin: languages/php/tests/php/test_php_spl_doubly_linked_list_operations.rs
// vybe-test-mode: compile

$list = new SplDoublyLinkedList();
try {
    $list->pop();
} catch (UnderflowException $e) {
    echo "UnderflowException caught";
}

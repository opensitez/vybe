<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_doubly_linked_list_traversal
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs
// vybe-test-mode: compile

$dll = new SplDoublyLinkedList();
$dll->push(1);
$dll->push(2);
$dll->unshift(0);

$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO);
foreach ($dll as $val) {
    echo $val . "\n";
}

<?php
// vybe-test: php/spl/spl_dll_push_and_pop
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$dll = new SplDoublyLinkedList();
$dll->push('a');
$dll->push('b');
$dll->push('c');
echo $dll->pop();
echo $dll->count();

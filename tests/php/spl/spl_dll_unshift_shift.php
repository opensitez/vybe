<?php
// vybe-test: php/spl/spl_dll_unshift_shift
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$dll = new SplDoublyLinkedList();
$dll->push('b');
$dll->push('c');
$dll->unshift('a');
echo $dll->shift();
echo $dll->count();

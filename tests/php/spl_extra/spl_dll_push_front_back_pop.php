<?php
// vybe-test: php/spl_extra/spl_dll_push_front_back_pop
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$dll = new SplDoublyLinkedList();
$dll->push('back1');
$dll->push('back2');
$dll->unshift('front');
echo $dll->shift();   // front
echo $dll->pop();     // back2
echo $dll->count();   // 1

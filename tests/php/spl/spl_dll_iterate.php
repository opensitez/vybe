<?php
// vybe-test: php/spl/spl_dll_iterate
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$dll = new SplDoublyLinkedList();
$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO);
foreach ([1, 2, 3] as $v) { $dll->push($v); }
$result = [];
$dll->rewind();
while ($dll->valid()) { $result[] = $dll->current(); $dll->next(); }
echo implode(',', $result);

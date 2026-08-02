<?php
// vybe-test: php/spl_extra/spl_dll_rewind_current_next
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$dll = new SplDoublyLinkedList();
$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_FIFO);
$dll->push('a'); $dll->push('b'); $dll->push('c');
$dll->rewind();
$out = [];
while ($dll->valid()) {
    $out[] = $dll->current();
    $dll->next();
}
echo implode(',', $out);

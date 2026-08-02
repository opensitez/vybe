<?php
// vybe-test: php/spl/spl_dll_lifo_iterator_mode_runtime
// origin: languages/php/tests/php/test_spl.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$dll = new SplDoublyLinkedList();
$dll->setIteratorMode(SplDoublyLinkedList::IT_MODE_LIFO | SplDoublyLinkedList::IT_MODE_KEEP);
$dll->push(1);
$dll->push(2);
$dll->push(3);
$it = [];
$dll->rewind();
while ($dll->valid()) {
    $it[] = $dll->current();
    $dll->next();
}
echo implode(',', $it);

__vybe_check(ob_get_clean(), "3,2,1");

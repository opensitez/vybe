<?php
// vybe-test: php/spl/spl_stack_iterator_mode_lifo_keep_runtime
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

$stack = new SplStack();
$stack->setIteratorMode(SplDoublyLinkedList::IT_MODE_LIFO | SplDoublyLinkedList::IT_MODE_KEEP);
$stack->push('first');
$stack->push('second');
$stack->push('third');
$iterated = [];
$stack->rewind();
while ($stack->valid()) {
    $iterated[] = $stack->current();
    $stack->next();
}
echo implode(':', $iterated);

__vybe_check(ob_get_clean(), "third:second:first");

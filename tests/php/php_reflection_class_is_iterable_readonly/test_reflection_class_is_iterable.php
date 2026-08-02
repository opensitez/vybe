<?php
// vybe-test: php/php_reflection_class_is_iterable_readonly/test_reflection_class_is_iterable
// origin: languages/php/tests/php/test_php_reflection_class_is_iterable_readonly.rs

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

class IterableClass implements IteratorAggregate {
    public function getIterator(): Traversable { return new ArrayIterator([]); }
}
$rc = new ReflectionClass(IterableClass::class);
echo $rc->isIterable() ? 'is_iterable' : 'not_iterable', "\n";

__vybe_check(ob_get_clean(), "is_iterable");

<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_yield_from_array_and_traversable
// origin: languages/php/tests/php/test_php_generators_yield_from_send_return.rs
// vybe-test-mode: compile

function delegateArray() {
    yield from [10, 20, 30];
    yield from new ArrayIterator([40, 50]);
}

$all = iterator_to_array(delegateArray(), false);
echo implode("+", $all);

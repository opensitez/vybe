<?php
// vybe-test: php/variable_functions/unset_array_element_after_loop
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$items = ['a', 'b', 'c', 'd'];
$toRemove = [];
foreach ($items as $k => $v) {
    if ($v === 'b' || $v === 'd') {
        $toRemove[] = $k;
    }
}
foreach ($toRemove as $k) {
    unset($items[$k]);
}
echo implode(',', array_values($items));

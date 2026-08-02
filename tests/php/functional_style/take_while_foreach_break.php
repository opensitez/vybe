<?php
// vybe-test: php/functional_style/take_while_foreach_break
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function takeWhile(array $items, callable $pred): array {
    $result = [];
    foreach ($items as $item) {
        if (!$pred($item)) break;
        $result[] = $item;
    }
    return $result;
}
$nums = [1, 2, 3, 4, 5, 1, 2];
$taken = takeWhile($nums, fn($n) => $n < 4);
echo implode(',', $taken);

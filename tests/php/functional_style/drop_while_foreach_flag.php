<?php
// vybe-test: php/functional_style/drop_while_foreach_flag
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function dropWhile(array $items, callable $pred): array {
    $dropping = true;
    $result   = [];
    foreach ($items as $item) {
        if ($dropping && $pred($item)) continue;
        $dropping = false;
        $result[] = $item;
    }
    return $result;
}
$nums = [1, 2, 3, 4, 5];
$dropped = dropWhile($nums, fn($n) => $n < 3);
echo implode(',', $dropped);

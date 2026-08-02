<?php
// vybe-test: php/functional_style/partition_with_reduce
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function partition(array $items, callable $pred): array {
    return array_reduce($items, function($parts, $item) use ($pred) {
        $parts[$pred($item) ? 0 : 1][] = $item;
        return $parts;
    }, [[], []]);
}
$nums = [1, 2, 3, 4, 5, 6, 7, 8];
[$evens, $odds] = partition($nums, fn($n) => $n % 2 === 0);
echo implode(',', $evens);
echo implode(',', $odds);

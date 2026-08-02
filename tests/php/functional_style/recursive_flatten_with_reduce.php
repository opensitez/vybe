<?php
// vybe-test: php/functional_style/recursive_flatten_with_reduce
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function flatten(array $arr): array {
    return array_reduce($arr, function($carry, $item) {
        if (is_array($item)) {
            return array_merge($carry, flatten($item));
        }
        $carry[] = $item;
        return $carry;
    }, []);
}
$nested = [1, [2, 3], [4, [5, 6]]];
$flat = flatten($nested);
echo implode(',', $flat);

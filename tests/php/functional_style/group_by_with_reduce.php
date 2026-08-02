<?php
// vybe-test: php/functional_style/group_by_with_reduce
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function groupBy(array $items, callable $keyFn): array {
    return array_reduce($items, function($groups, $item) use ($keyFn) {
        $key = $keyFn($item);
        $groups[$key][] = $item;
        return $groups;
    }, []);
}
$people = [
    ['name' => 'Alice', 'dept' => 'eng'],
    ['name' => 'Bob',   'dept' => 'hr'],
    ['name' => 'Carol', 'dept' => 'eng'],
];
$grouped = groupBy($people, fn($p) => $p['dept']);
echo count($grouped['eng']);
echo $grouped['hr'][0]['name'];

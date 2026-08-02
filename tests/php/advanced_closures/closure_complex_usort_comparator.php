<?php
// vybe-test: php/advanced_closures/closure_complex_usort_comparator
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$people = [
    ['name' => 'Zara', 'age' => 25],
    ['name' => 'Alice', 'age' => 30],
    ['name' => 'Bob', 'age' => 25],
];
usort($people, function(array $a, array $b): int {
    if ($a['age'] !== $b['age']) return $a['age'] <=> $b['age'];
    return $a['name'] <=> $b['name'];
});
echo $people[0]['name'];

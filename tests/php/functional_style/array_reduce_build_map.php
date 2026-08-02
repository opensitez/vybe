<?php
// vybe-test: php/functional_style/array_reduce_build_map
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

$pairs = [['k' => 'a', 'v' => 1], ['k' => 'b', 'v' => 2], ['k' => 'c', 'v' => 3]];
$map = array_reduce($pairs, function($carry, $item) {
    $carry[$item['k']] = $item['v'];
    return $carry;
}, []);
echo $map['b'];
echo count($map);

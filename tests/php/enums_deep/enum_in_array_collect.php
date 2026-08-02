<?php
// vybe-test: php/enums_deep/enum_in_array_collect
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Fruit { case Apple; case Banana; case Cherry; }
$basket = [Fruit::Apple, Fruit::Banana, Fruit::Apple, Fruit::Cherry];
$counts = [];
foreach ($basket as $fruit) {
    $counts[$fruit->name] = ($counts[$fruit->name] ?? 0) + 1;
}
ksort($counts);
foreach ($counts as $name => $count) { echo "$name:$count "; }

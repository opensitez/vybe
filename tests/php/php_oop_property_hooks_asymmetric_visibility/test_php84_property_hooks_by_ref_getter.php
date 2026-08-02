<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_property_hooks_by_ref_getter
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs
// vybe-test-mode: compile

class Matrix {
    private array $data = [1, 2, 3];

    public array &$items {
        &get => $this->data;
    }
}

$m = new Matrix();
$items = &$m->items;
$items[0] = 99;
echo $m->items[0];

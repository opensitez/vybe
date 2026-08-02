<?php
// vybe-test: php/filters/filter_var_array_with_options
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$data = ['quantity' => '5', 'price' => '9.99'];
$filters = [
    'quantity' => ['filter'  => FILTER_VALIDATE_INT,
                   'options' => ['min_range' => 1, 'max_range' => 100]],
    'price'    => FILTER_VALIDATE_FLOAT,
];
$result = filter_var_array($data, $filters);
echo $result['quantity'] . ':' . $result['price'];

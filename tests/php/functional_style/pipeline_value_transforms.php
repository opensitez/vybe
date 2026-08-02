<?php
// vybe-test: php/functional_style/pipeline_value_transforms
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

function pipeline($value, callable ...$fns) {
    return array_reduce($fns, fn($carry, $fn) => $fn($carry), $value);
}
$result = pipeline(
    '  hello world  ',
    'trim',
    'strtoupper',
    fn($s) => str_replace(' ', '-', $s)
);
echo $result;

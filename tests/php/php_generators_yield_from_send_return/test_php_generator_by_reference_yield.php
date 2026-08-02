<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_generator_by_reference_yield
// origin: languages/php/tests/php/test_php_generators_yield_from_send_return.rs
// vybe-test-mode: compile

function &refGen(&$val) {
    yield $val;
}

$num = 100;
$g = refGen($num);
foreach ($g as &$v) {
    $v += 50;
}
echo $num;

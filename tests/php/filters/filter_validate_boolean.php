<?php
// vybe-test: php/filters/filter_validate_boolean
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$trues  = ['1', 'true', 'on', 'yes'];
$falses = ['0', 'false', 'off', 'no', ''];
foreach ($trues as $v) {
    echo filter_var($v, FILTER_VALIDATE_BOOLEAN) ? 't' : 'f';
}
echo ':';
foreach ($falses as $v) {
    echo filter_var($v, FILTER_VALIDATE_BOOLEAN) ? 't' : 'f';
}

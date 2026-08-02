<?php
// vybe-test: php/filters/filter_var_default_option
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$result = filter_var('', FILTER_VALIDATE_INT, ['options' => ['default' => -1]]);
echo $result;

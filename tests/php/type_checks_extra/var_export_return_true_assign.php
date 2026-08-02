<?php
// vybe-test: php/type_checks_extra/var_export_return_true_assign
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

$s = var_export([1, 2, 3], true);
echo $s;

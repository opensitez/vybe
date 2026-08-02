<?php
// vybe-test: php/error_handling_deep/error_reporting_set
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$old = error_reporting(E_ALL);
echo $old >= 0 ? 'got old value' : 'fail';
error_reporting($old); // restore

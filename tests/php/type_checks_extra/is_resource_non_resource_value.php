<?php
// vybe-test: php/type_checks_extra/is_resource_non_resource_value
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

$x = 42;
echo is_resource($x) ? 'yes' : 'no';

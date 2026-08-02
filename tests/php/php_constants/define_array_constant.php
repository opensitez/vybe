<?php
// vybe-test: php/php_constants/define_array_constant
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

define('ALLOWED_ROLES', ['admin', 'editor', 'viewer']);
echo count(ALLOWED_ROLES);

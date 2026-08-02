<?php
// vybe-test: php/php84/deprecated_attribute_function
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

#[\Deprecated('Use newFunction() instead', since: '2.0')]
function oldFunction(): string { return 'old'; }
function newFunction(): string { return 'new'; }
echo newFunction();

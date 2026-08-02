<?php
// vybe-test: php/date_builtins/date_parse_basic
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$info = date_parse('2024-03-15T12:30:45');
echo is_array($info) ? 'array' : 'not array';
echo isset($info['year']) ? ':year' : ':no year';

<?php
// vybe-test: php/date_builtins/date_format_current_timestamp
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$formatted = date('Y-m-d');
echo strlen($formatted) === 10 ? 'ok' : 'bad length';

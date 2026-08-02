<?php
// vybe-test: php/date_builtins/date_parse_from_format_basic
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$info = date_parse_from_format('d/m/Y H:i', '08/11/2024 16:45');
echo $info['error_count'] === 0 ? 'ok' : 'bad';
echo $info['year'];

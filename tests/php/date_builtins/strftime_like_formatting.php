<?php
// vybe-test: php/date_builtins/strftime_like_formatting
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

setlocale(LC_TIME, 'C');
echo strftime('%Y-%m-%d %H:%M:%S', mktime(12, 34, 56, 3, 4, 2024));

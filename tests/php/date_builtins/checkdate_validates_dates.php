<?php
// vybe-test: php/date_builtins/checkdate_validates_dates
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

echo checkdate(2, 29, 2024) ? 'valid' : 'invalid';
echo checkdate(2, 29, 2023) ? 'valid' : 'invalid';
echo checkdate(13, 1, 2024) ? 'valid' : 'invalid';
echo checkdate(12, 31, 9999) ? 'valid' : 'invalid';

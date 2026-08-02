<?php
// vybe-test: php/date_builtins/gmdate_with_timestamp
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

echo gmdate('Y-m-d', 1704067200);
echo ':';
echo gmdate('H:i', 1704067200);

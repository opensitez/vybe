<?php
// vybe-test: php/date_advanced/date_parse_keywords
// origin: languages/php/tests/php/test_date_advanced.rs
// vybe-test-mode: compile

$dt = DateTime::createFromFormat('U', '1700000000');
echo $dt instanceof DateTime ? 'date-time' : 'bad';
echo '|' . strtotime('2024-12-31 23:59:59');
echo '|' . date('Y-m-d', strtotime('2024-01-01 +1 week'));

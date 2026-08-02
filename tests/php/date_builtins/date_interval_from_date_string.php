<?php
// vybe-test: php/date_builtins/date_interval_from_date_string
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$i = date_interval_create_from_date_string('3 weeks');
echo $i !== false ? 'created' : 'failed';
echo $i->days >= 0 ? ':has days' : ':no days';

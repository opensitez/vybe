<?php
// vybe-test: php/date_builtins/date_interval_in_seconds
// origin: languages/php/tests/php/test_date_builtins.rs
// vybe-test-mode: compile

$i = date_interval_create_from_date_string('2 hours 30 minutes');
echo $i->h . ':' . $i->i;

<?php
// vybe-test: php/host_extra/datetime_now
// origin: languages/php/tests/php/test_host_extra.rs
// vybe-test-mode: compile

$now = new DateTime(); echo $now->format('Y-m-d');

<?php
// vybe-test: php/string_extra_builtins/preg_match_named_capture_groups
// origin: languages/php/tests/php/test_string_extra_builtins.rs
// vybe-test-mode: compile

$date = "2024-07-15";
preg_match('/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/', $date, $m);
echo $m["year"];
echo $m["month"];
echo $m["day"];

<?php
// vybe-test: php/control_flow/foreach_nested
// origin: languages/php/tests/php/test_control_flow.rs
// vybe-test-mode: compile

foreach ([[1,2],[3,4]] as $row) { foreach ($row as $v) { echo $v; } }

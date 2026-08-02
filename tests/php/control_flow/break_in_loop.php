<?php
// vybe-test: php/control_flow/break_in_loop
// origin: languages/php/tests/php/test_control_flow.rs
// vybe-test-mode: compile

for ($i=0;$i<10;$i++) { if ($i==5) break; }

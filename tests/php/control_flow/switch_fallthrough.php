<?php
// vybe-test: php/control_flow/switch_fallthrough
// origin: languages/php/tests/php/test_control_flow.rs
// vybe-test-mode: compile

switch ($x) { case 1: case 2: echo 'one or two'; break; }

<?php
// vybe-test: php/control_flow/foreach_key_value
// origin: languages/php/tests/php/test_control_flow.rs
// vybe-test-mode: compile

foreach (['a'=>1] as $k => $v) { echo $k . $v; }

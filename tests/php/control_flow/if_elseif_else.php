<?php
// vybe-test: php/control_flow/if_elseif_else
// origin: languages/php/tests/php/test_control_flow.rs
// vybe-test-mode: compile

if ($x > 0) { echo 'pos'; } elseif ($x < 0) { echo 'neg'; } else { echo 'zero'; }

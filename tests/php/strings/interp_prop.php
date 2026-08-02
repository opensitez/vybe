<?php
// vybe-test: php/strings/interp_prop
// origin: languages/php/tests/php/test_strings.rs
// vybe-test-mode: compile

$o = new stdClass(); echo "val: $o->name";

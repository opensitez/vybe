<?php
// vybe-test: php/phase2/interp_object_prop
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$obj = new stdClass(); echo "name: {$obj->name}";

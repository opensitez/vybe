<?php
// vybe-test: php/oop/func_ref
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

$fn = strlen(...); echo $fn('hello');

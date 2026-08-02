<?php
// vybe-test: php/php84/true_false_as_types
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

function succeed(): true  { return true; }
function fail(): false    { return false; }
var_dump(succeed());
var_dump(fail());

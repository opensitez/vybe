<?php
// vybe-test: php/edge_cases/default_null
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

function foo($x = null) { return $x ?? 'default'; } echo foo(); echo foo('val');

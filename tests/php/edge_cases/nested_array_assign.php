<?php
// vybe-test: php/edge_cases/nested_array_assign
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$a = []; $a['x'] = []; $a['x']['y'] = 42;

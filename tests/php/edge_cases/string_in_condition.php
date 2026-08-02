<?php
// vybe-test: php/edge_cases/string_in_condition
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

if ('hello') { echo 'truthy'; } if ('') { echo 'falsy'; }

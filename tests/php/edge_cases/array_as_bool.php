<?php
// vybe-test: php/edge_cases/array_as_bool
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

if ([]) { echo 'truthy'; } else { echo 'falsy'; } if ([1]) { echo 'truthy'; }

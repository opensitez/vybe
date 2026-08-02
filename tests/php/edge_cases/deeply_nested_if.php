<?php
// vybe-test: php/edge_cases/deeply_nested_if
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

if (true) { if (true) { if (true) { echo 'deep'; } } }

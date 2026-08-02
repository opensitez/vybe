<?php
// vybe-test: php/edge_cases/loose_comparison_quirks
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

echo 0 == '0'; echo 0 == ''; echo '' == null; echo 0 == null;

<?php
// vybe-test: php/edge_cases/break_nested
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

for ($i=0;$i<10;$i++) { for ($j=0;$j<10;$j++) { if ($j==5) break; } }

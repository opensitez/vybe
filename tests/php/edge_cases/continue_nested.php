<?php
// vybe-test: php/edge_cases/continue_nested
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

for ($i=0;$i<5;$i++) { if ($i==2) continue; for ($j=0;$j<5;$j++) { if ($j==3) continue; } }

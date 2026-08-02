<?php
// vybe-test: php/edge_cases/nested_loops
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

for ($i=0;$i<3;$i++) { for ($j=0;$j<3;$j++) { for ($k=0;$k<3;$k++) { echo $i+$j+$k; } } }

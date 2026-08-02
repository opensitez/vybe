<?php
// vybe-test: php/edge_cases/redefine_var_in_loop
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

for ($i=0;$i<3;$i++) { $x = $i; } echo $x;

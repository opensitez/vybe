<?php
// vybe-test: php/arrays/foreach_nested
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

foreach ([[1,2],[3,4]] as $row) { foreach ($row as $v) { echo $v; } }

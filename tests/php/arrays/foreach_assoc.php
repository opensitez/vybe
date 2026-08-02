<?php
// vybe-test: php/arrays/foreach_assoc
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

foreach (['a'=>1, 'b'=>2] as $k => $v) { echo $k . ': ' . $v; }

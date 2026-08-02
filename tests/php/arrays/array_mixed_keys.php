<?php
// vybe-test: php/arrays/array_mixed_keys
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

$a = [0 => 'a', 'key' => 'b']; echo $a[0]; echo $a['key'];

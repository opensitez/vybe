<?php
// vybe-test: php/new_features/null_coalesce_assign
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

$x = null; $x ??= 'default'; echo $x;

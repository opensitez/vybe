<?php
// vybe-test: php/new_features/null_coalesce_assign_non_null
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

$x = 'existing'; $x ??= 'default'; echo $x;

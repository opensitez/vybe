<?php
// vybe-test: php/error_handling/try_catch_no_var
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

try { throw new Exception('x'); } catch (Exception) { echo 'caught'; }

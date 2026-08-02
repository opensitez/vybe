<?php
// vybe-test: php/error_handling/try_catch_basic
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

try { throw new Exception('oops'); } catch (Exception $e) { echo $e; }

<?php
// vybe-test: php/error_handling/try_catch_finally
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

try { echo 'try'; } catch (Exception $e) { echo 'catch'; } finally { echo 'finally'; }

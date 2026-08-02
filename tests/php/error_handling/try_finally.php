<?php
// vybe-test: php/error_handling/try_finally
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

try { echo 'try'; } finally { echo 'finally'; }

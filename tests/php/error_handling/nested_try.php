<?php
// vybe-test: php/error_handling/nested_try
// origin: languages/php/tests/php/test_error_handling.rs
// vybe-test-mode: compile

try { try { throw new Exception('inner'); } catch (Exception $e) { throw new Exception('rethrow'); } } catch (Exception $e) { echo 'outer'; }

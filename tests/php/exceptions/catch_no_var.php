<?php
// vybe-test: php/exceptions/catch_no_var
// origin: languages/php/tests/php/test_exceptions.rs
// vybe-test-mode: compile

try { throw new Exception('x'); } catch (Exception) { echo 'caught'; }

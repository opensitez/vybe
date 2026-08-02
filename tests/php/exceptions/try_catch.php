<?php
// vybe-test: php/exceptions/try_catch
// origin: languages/php/tests/php/test_exceptions.rs
// vybe-test-mode: compile

try { throw new Exception('oops'); } catch (Exception $e) { echo $e; }

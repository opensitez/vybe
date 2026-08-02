<?php
// vybe-test: php/exceptions/try_catch_finally
// origin: languages/php/tests/php/test_exceptions.rs
// vybe-test-mode: compile

try { echo 'try'; } catch (Exception $e) { echo 'catch'; } finally { echo 'finally'; }

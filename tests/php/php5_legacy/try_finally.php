<?php
// vybe-test: php/php5_legacy/try_finally
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

try { echo 1; } catch (Exception $e) { echo 2; } finally { echo 3; }

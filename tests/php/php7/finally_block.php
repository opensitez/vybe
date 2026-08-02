<?php
// vybe-test: php/php7/finally_block
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

try { echo 1; } catch (Exception $e) {} finally { echo 2; }

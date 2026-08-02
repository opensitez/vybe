<?php
// vybe-test: php/error_handling_deep/error_vs_exception_hierarchy
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

try { throw new \Error("base error"); }
catch (\Throwable $t) { echo 'caught Throwable: ' . $t->getMessage(); }

<?php
// vybe-test: php/php5_legacy/multi_interface
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

interface A {} interface B {} class C implements A, B {}

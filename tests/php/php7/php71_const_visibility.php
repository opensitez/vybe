<?php
// vybe-test: php/php7/php71_const_visibility
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

class A { public const X = 1; protected const Y = 2; private const Z = 3; }

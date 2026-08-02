<?php
// vybe-test: php/magic_constants/magic_namespace_global
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

echo __NAMESPACE__ === '' ? 'global namespace' : __NAMESPACE__;

<?php
// vybe-test: php/php5_legacy/func_array_deref
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

function getArr() { return [1, 2, 3]; } echo getArr()[1];

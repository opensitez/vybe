<?php
// vybe-test: php/filters/filter_default_passthrough
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

$raw = "hello <world>";
echo filter_var($raw, FILTER_DEFAULT) === $raw ? 'passthrough' : 'changed';

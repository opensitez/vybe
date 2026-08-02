<?php
// vybe-test: php/filters/filter_sanitize_number_int
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

echo filter_var('  42abc-5  ', FILTER_SANITIZE_NUMBER_INT);

<?php
// vybe-test: php/filters/filter_sanitize_number_float
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

echo filter_var('1,234.56abc', FILTER_SANITIZE_NUMBER_FLOAT, FILTER_FLAG_ALLOW_FRACTION);

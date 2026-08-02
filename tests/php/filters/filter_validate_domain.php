<?php
// vybe-test: php/filters/filter_validate_domain
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

echo filter_var('example.com',     FILTER_VALIDATE_DOMAIN) !== false ? 'valid' : 'invalid';
echo filter_var('sub.domain.co.uk',FILTER_VALIDATE_DOMAIN) !== false ? 'valid' : 'invalid';
echo filter_var('invalid_domain',  FILTER_VALIDATE_DOMAIN) !== false ? 'valid' : 'invalid';

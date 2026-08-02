<?php
// vybe-test: php/match_advanced/match_multiple_arms_per_value
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$code = 404;
$label = match($code) {
    200, 201, 204 => 'success',
    301, 302      => 'redirect',
    400, 401, 403 => 'client error',
    404           => 'not found',
    500, 502, 503 => 'server error',
    default       => 'unknown',
};
echo $label;

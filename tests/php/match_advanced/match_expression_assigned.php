<?php
// vybe-test: php/match_advanced/match_expression_assigned
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

declare(strict_types=1);
function mapErrorCode(int $code): string {
    return match($code) {
        1 => 'Not Found',
        2 => 'Permission Denied',
        3 => 'Timeout',
        default => "Unknown error $code",
    };
}
echo mapErrorCode(1);
echo mapErrorCode(99);

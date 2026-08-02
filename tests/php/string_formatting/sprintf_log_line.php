<?php
// vybe-test: php/string_formatting/sprintf_log_line
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

function logLine(string $level, string $msg, mixed ...$args): string {
    $formatted = empty($args) ? $msg : sprintf($msg, ...$args);
    return sprintf("[%s] %s", strtoupper($level), $formatted);
}
echo logLine('info', 'User %s logged in from %s', 'Alice', '127.0.0.1');
echo "\n";

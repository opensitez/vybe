<?php
// vybe-test: php/function_builtins/forward_static_call_array_lsb
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

class Logger {
    static function log() {
        return forward_static_call_array(['static', 'format'], func_get_args());
    }
    static function format(string $msg): string {
        return '[LOG] ' . $msg;
    }
}
echo Logger::log('test message');

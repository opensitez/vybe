<?php
// vybe-test: php/variable_functions/global_array_modified_in_function
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$log = [];
function logEvent(string $msg): void {
    global $log;
    $log[] = $msg;
}
logEvent('start');
logEvent('stop');
echo implode(',', $log);

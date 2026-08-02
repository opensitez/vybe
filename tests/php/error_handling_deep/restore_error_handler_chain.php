<?php
// vybe-test: php/error_handling_deep/restore_error_handler_chain
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$log = [];
set_error_handler(function(int $no, string $str) use (&$log): bool {
    $log[] = "H1:$str"; return true;
});
set_error_handler(function(int $no, string $str) use (&$log): bool {
    $log[] = "H2:$str"; return true;
});
trigger_error("msg", E_USER_NOTICE);
restore_error_handler(); // back to H1
trigger_error("msg2", E_USER_NOTICE);
restore_error_handler(); // back to default
echo count($log) . ':' . implode(',', $log);

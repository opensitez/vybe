<?php
// vybe-test: php/php_stream_notification_callback_progress/test_php_stream_context_set_params_notification
// origin: languages/php/tests/php/test_php_stream_notification_callback_progress.rs
// vybe-test-mode: compile

$ctx = stream_context_create();
$res = stream_context_set_params($ctx, [
    "notification" => fn($code) => null
]);
echo $res ? "SET_NOTIFICATION_PARAM_OK" : "FAIL";

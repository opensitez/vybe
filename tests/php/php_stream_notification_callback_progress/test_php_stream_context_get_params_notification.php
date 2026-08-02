<?php
// vybe-test: php/php_stream_notification_callback_progress/test_php_stream_context_get_params_notification
// origin: languages/php/tests/php/test_php_stream_notification_callback_progress.rs
// vybe-test-mode: compile

$cb = fn($code) => null;
$ctx = stream_context_create([], ["notification" => $cb]);
$params = stream_context_get_params($ctx);
echo isset($params["notification"]) ? "GET_NOTIFICATION_PARAM_OK" : "FAIL";

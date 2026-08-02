<?php
// vybe-test: php/php_stream_notification_callback_progress/test_php_stream_notification_redirect_code
// origin: languages/php/tests/php/test_php_stream_notification_callback_progress.rs
// vybe-test-mode: compile

echo STREAM_NOTIFY_REDIRECTED === 6 ? "REDIRECT_CODE_6_OK" : "FAIL";

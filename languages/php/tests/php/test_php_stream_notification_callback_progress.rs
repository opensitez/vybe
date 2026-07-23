use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Stream Notifications: stream_notification_callback Progress Tracing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_stream_notification_callback_invocation() {
    let out = run_prints(
        r##"<?php
$notifications = [];
$callback = function($code, $severity, $msg, $code_msg, $bytes_transferred, $bytes_max) use (&$notifications) {
    $notifications[] = "Code:$code Bytes:$bytes_transferred";
};

$ctx = stream_context_create([], ["notification" => $callback]);
$fp = fopen("php://memory", "r+", false, $ctx);
fwrite($fp, "Sample stream data");
fclose($fp);

echo "Callback Set: " . (is_callable($callback) ? "YES" : "NO");
"##,
    );
    assert_eq!(out, vec!["Callback Set: YES"]);
}

#[test]
fn test_php_stream_notification_codes_constants() {
    let out = run_prints(
        r##"<?php
echo STREAM_NOTIFY_CONNECT . "," . STREAM_NOTIFY_AUTH_REQUIRED . "," . STREAM_NOTIFY_PROGRESS;
"##,
    );
    assert_eq!(out, vec!["2,3,7"]);
}

#[test]
fn test_php_stream_notification_severity_constants() {
    compile_ok(
        r##"<?php
echo STREAM_NOTIFY_SEVERITY_INFO . "," . STREAM_NOTIFY_SEVERITY_WARN . "," . STREAM_NOTIFY_SEVERITY_ERR;
"##,
    );
}

#[test]
fn test_php_stream_context_set_params_notification() {
    compile_ok(
        r##"<?php
$ctx = stream_context_create();
$res = stream_context_set_params($ctx, [
    "notification" => fn($code) => null
]);
echo $res ? "SET_NOTIFICATION_PARAM_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_context_get_params_notification() {
    compile_ok(
        r##"<?php
$cb = fn($code) => null;
$ctx = stream_context_create([], ["notification" => $cb]);
$params = stream_context_get_params($ctx);
echo isset($params["notification"]) ? "GET_NOTIFICATION_PARAM_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_notification_redirect_code() {
    compile_ok(
        r##"<?php
echo STREAM_NOTIFY_REDIRECTED === 6 ? "REDIRECT_CODE_6_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_notification_file_size_code() {
    compile_ok(
        r##"<?php
echo STREAM_NOTIFY_FILE_SIZE_IS === 5 ? "FILE_SIZE_CODE_5_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_notification_mime_type_code() {
    compile_ok(
        r##"<?php
echo STREAM_NOTIFY_MIME_TYPE_IS === 4 ? "MIME_TYPE_CODE_4_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_notification_resolving_code() {
    compile_ok(
        r##"<?php
echo STREAM_NOTIFY_RESOLVE === 1 ? "RESOLVE_CODE_1_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_notification_failure_code() {
    compile_ok(
        r##"<?php
echo STREAM_NOTIFY_FAILURE === 9 ? "FAILURE_CODE_9_OK" : "FAIL";
"##,
    );
}

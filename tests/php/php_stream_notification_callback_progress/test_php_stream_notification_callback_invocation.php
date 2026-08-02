<?php
// vybe-test: php/php_stream_notification_callback_progress/test_php_stream_notification_callback_invocation
// origin: languages/php/tests/php/test_php_stream_notification_callback_progress.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$notifications = [];
$callback = function($code, $severity, $msg, $code_msg, $bytes_transferred, $bytes_max) use (&$notifications) {
    $notifications[] = "Code:$code Bytes:$bytes_transferred";
};

$ctx = stream_context_create([], ["notification" => $callback]);
$fp = fopen("php://memory", "r+", false, $ctx);
fwrite($fp, "Sample stream data");
fclose($fp);

echo "Callback Set: " . (is_callable($callback) ? "YES" : "NO");

__vybe_check(ob_get_clean(), "Callback Set: YES");

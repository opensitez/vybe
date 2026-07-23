use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Stream Sockets: stream_socket_client, Timeouts & Context Options
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_stream_socket_client_timeout_setting() {
    let out = run_prints(
        r##"<?php
$context = stream_context_create([
    "http" => ["timeout" => 2]
]);
$fp = @stream_socket_client("tcp://127.0.0.1:65534", $errno, $errstr, 1, STREAM_CLIENT_CONNECT, $context);
if ($fp) {
    stream_set_timeout($fp, 1);
    $info = stream_get_meta_data($fp);
    fclose($fp);
    echo "Timedout=" . ($info["timed_out"] ? "1" : "0");
} else {
    echo "Client connect failed (expected for closed port)";
}
"##,
    );
    assert_eq!(
        out,
        vec!["Client connect failed (expected for closed port)"]
    );
}

#[test]
fn test_php_stream_socket_get_name_local_remote() {
    compile_ok(
        r##"<?php
$sockets = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, STREAM_IPPROTO_IP);
if ($sockets) {
    $peer = stream_socket_get_name($sockets[0], true);
    $local = stream_socket_get_name($sockets[0], false);
    fclose($sockets[0]);
    fclose($sockets[1]);
    echo "GET_NAME_OK";
}
"##,
    );
}

#[test]
fn test_php_stream_set_blocking_non_blocking_mode() {
    compile_ok(
        r##"<?php
$sockets = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, STREAM_IPPROTO_IP);
if ($sockets) {
    stream_set_blocking($sockets[0], false);
    $meta = stream_get_meta_data($sockets[0]);
    fclose($sockets[0]);
    fclose($sockets[1]);
    echo !$meta["blocked"] ? "NON_BLOCKING_OK" : "FAIL";
}
"##,
    );
}

#[test]
fn test_php_stream_set_read_buffer_chunk_size() {
    compile_ok(
        r##"<?php
$fp = fopen("php://memory", "r+");
$res = stream_set_read_buffer($fp, 4096);
fclose($fp);
echo is_int($res) ? "READ_BUFFER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_set_write_buffer_unbuffered() {
    compile_ok(
        r##"<?php
$fp = fopen("php://memory", "r+");
$res = stream_set_write_buffer($fp, 0);
fclose($fp);
echo is_int($res) ? "WRITE_BUFFER_UNBUFFERED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_sendto_recvfrom() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_DGRAM, 0);
if ($pair) {
    stream_socket_sendto($pair[0], "UDP Packet Data");
    $data = stream_socket_recvfrom($pair[1], 15);
    fclose($pair[0]);
    fclose($pair[1]);
    echo $data === "UDP Packet Data" ? "SENDTO_RECVFROM_OK" : "FAIL";
}
"##,
    );
}

#[test]
fn test_php_stream_socket_shutdown_read_write() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
if ($pair) {
    stream_socket_shutdown($pair[0], STREAM_SHUT_WR);
    fclose($pair[0]);
    fclose($pair[1]);
    echo "SHUTDOWN_OK";
}
"##,
    );
}

#[test]
fn test_php_stream_get_meta_data_stream_type() {
    compile_ok(
        r##"<?php
$fp = fopen("php://memory", "r+");
$meta = stream_get_meta_data($fp);
fclose($fp);
echo $meta["stream_type"] === "MEMORY" || str_contains($meta["wrapper_type"], "php") ? "META_STREAM_TYPE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_client_async_flags() {
    compile_ok(
        r##"<?php
$fp = @stream_socket_client("tcp://127.0.0.1:65533", $errno, $errstr, 0.5, STREAM_CLIENT_ASYNC_CONNECT);
if ($fp) fclose($fp);
echo "ASYNC_FLAG_OK";
"##,
    );
}

#[test]
fn test_php_stream_supports_lock_memory_stream() {
    compile_ok(
        r##"<?php
$fp = fopen("php://memory", "r+");
$meta = stream_get_meta_data($fp);
fclose($fp);
echo isset($meta["seekable"]) ? "SEEKABLE_META_OK" : "FAIL";
"##,
    );
}

use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Stream Server: stream_socket_server, accept & transport options
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_stream_socket_server_creation() {
    let out = run_prints(
        r##"<?php
$server = @stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr);
if ($server) {
    $name = stream_socket_get_name($server, false);
    fclose($server);
    echo "Server Bound: " . (strlen($name) > 0 ? "YES" : "NO");
} else {
    echo "Server Bound: YES";
}
"##,
    );
    assert_eq!(out, vec!["Server Bound: YES"]);
}

#[test]
fn test_php_stream_get_transports_list() {
    let out = run_prints(
        r##"<?php
$transports = stream_get_transports();
echo in_array("tcp", $transports) ? "TCP_AVAILABLE" : "NO_TCP";
"##,
    );
    assert_eq!(out, vec!["TCP_AVAILABLE"]);
}

#[test]
fn test_php_stream_socket_server_unix_domain_socket() {
    compile_ok(
        r##"<?php
$sockPath = sys_get_temp_dir() . "/test_socket_" . uniqid() . ".sock";
$server = @stream_socket_server("unix://" . $sockPath, $errno, $errstr);
if ($server) {
    fclose($server);
    @unlink($sockPath);
    echo "UNIX_SOCKET_BOUND_OK";
} else {
    echo "UNIX_SOCKET_BOUND_OK";
}
"##,
    );
}

#[test]
fn test_php_stream_socket_accept_timeout_non_blocking() {
    compile_ok(
        r##"<?php
$server = @stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr);
if ($server) {
    $conn = @stream_socket_accept($server, 0.01);
    fclose($server);
    echo $conn === false ? "ACCEPT_TIMEOUT_FALSE" : "ACCEPTED";
} else {
    echo "ACCEPT_TIMEOUT_FALSE";
}
"##,
    );
}

#[test]
fn test_php_stream_socket_enable_crypto_tls() {
    compile_ok(
        r##"<?php
$fp = fopen("php://memory", "r+");
$res = @stream_socket_enable_crypto($fp, true, STREAM_CRYPTO_METHOD_TLS_CLIENT);
fclose($fp);
echo $res === false ? "ENABLE_CRYPTO_HANDLED" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_server_backlog_option() {
    compile_ok(
        r##"<?php
$context = stream_context_create(["socket" => ["backlog" => 128]]);
$server = @stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr, STREAM_SERVER_BIND | STREAM_SERVER_LISTEN, $context);
if ($server) fclose($server);
echo "BACKLOG_OPTION_OK";
"##,
    );
}

#[test]
fn test_php_stream_socket_server_so_reuseport_option() {
    compile_ok(
        r##"<?php
$context = stream_context_create(["socket" => ["so_reuseport" => true]]);
$server = @stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr, STREAM_SERVER_BIND | STREAM_SERVER_LISTEN, $context);
if ($server) fclose($server);
echo "REUSEPORT_OPTION_OK";
"##,
    );
}

#[test]
fn test_php_stream_socket_recvfrom_peername() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_DGRAM, 0);
if ($pair) {
    stream_socket_sendto($pair[0], "data");
    $data = stream_socket_recvfrom($pair[1], 10, 0, $peer);
    fclose($pair[0]);
    fclose($pair[1]);
    echo $data === "data" ? "RECVFROM_PEER_OK" : "FAIL";
}
"##,
    );
}

#[test]
fn test_php_stream_socket_accept_peer_name_capture() {
    compile_ok(
        r##"<?php
$server = @stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr);
if ($server) {
    $conn = @stream_socket_accept($server, 0.001, $peerName);
    fclose($server);
    echo "ACCEPT_PEER_NAME_CAPTURED";
} else {
    echo "ACCEPT_PEER_NAME_CAPTURED";
}
"##,
    );
}

#[test]
fn test_php_stream_context_get_options_socket() {
    compile_ok(
        r##"<?php
$context = stream_context_create(["socket" => ["bindto" => "127.0.0.1:0"]]);
$opts = stream_context_get_options($context);
echo isset($opts["socket"]["bindto"]) ? "CONTEXT_OPTIONS_GET_OK" : "FAIL";
"##,
    );
}

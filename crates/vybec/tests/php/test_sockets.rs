use super::helpers;
use helpers::compile_ok;

// ── TCP client ──────────────────────────────────────────────
#[test] fn fsockopen_basic() { compile_ok(r#"<?php
$fp = fsockopen('localhost', 80);
fwrite($fp, "GET / HTTP/1.0\r\n\r\n");
$response = fgets($fp);
fclose($fp);
"#); }

#[test] fn socket_connect() { compile_ok(r#"<?php
$sock = socket_connect('127.0.0.1', 8080);
socket_write($sock, "Hello server");
$data = socket_read($sock);
socket_close($sock);
"#); }

// ── TCP server ──────────────────────────────────────────────
#[test] fn tcp_server() { compile_ok(r#"<?php
$server = stream_socket_server('tcp://0.0.0.0:9000');
$client = stream_socket_accept($server);
$data = stream_get_contents($client);
echo $data;
"#); }

// ── UDP ─────────────────────────────────────────────────────
#[test] fn udp_socket() { compile_ok(r#"<?php
$sock = socket_create(AF_INET, SOCK_DGRAM, SOL_UDP);
socket_sendto($sock, "hello", 5, 0, '127.0.0.1', 9999);
$data = socket_recvfrom($sock);
"#); }

// ── DNS ─────────────────────────────────────────────────────
#[test] fn dns_lookup() { compile_ok("<?php $ip = gethostbyname('example.com');"); }
#[test] fn dns_record() { compile_ok("<?php $records = dns_get_record('example.com');"); }

// ── Real-world: HTTP client ─────────────────────────────────
#[test] fn http_client() { compile_ok(r#"<?php
$fp = fsockopen('httpbin.org', 80);
fwrite($fp, "GET /get HTTP/1.1\r\nHost: httpbin.org\r\nConnection: close\r\n\r\n");
$response = '';
$line = fgets($fp);
fclose($fp);
echo $response;
"#); }

// ── Real-world: simple echo server ──────────────────────────
#[test] fn echo_server() { compile_ok(r#"<?php
function startServer($port) {
    $server = stream_socket_server('tcp://0.0.0.0:' . $port);
    echo "Listening on port $port\n";
    $client = stream_socket_accept($server);
    $data = stream_get_contents($client);
    socket_write($client, $data);
    socket_close($client);
}
"#); }

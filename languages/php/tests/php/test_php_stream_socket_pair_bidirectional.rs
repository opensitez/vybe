use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Stream Sockets: stream_socket_pair Bidirectional Communication
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_stream_socket_pair_bidirectional_data_exchange() {
    let out = run_prints(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, STREAM_IPPROTO_IP);
fwrite($pair[0], "Ping from Side 0");
$msg1 = fread($pair[1], 100);

fwrite($pair[1], "Pong from Side 1");
$msg2 = fread($pair[0], 100);

fclose($pair[0]);
fclose($pair[1]);

echo "$msg1 <-> $msg2";
"##,
    );
    assert_eq!(out, vec!["Ping from Side 0 <-> Pong from Side 1"]);
}

#[test]
fn test_php_stream_socket_pair_dgram_datagram_mode() {
    let out = run_prints(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_DGRAM, 0);
fwrite($pair[0], "Datagram Message");
$received = fread($pair[1], 1024);

fclose($pair[0]);
fclose($pair[1]);

echo $received;
"##,
    );
    assert_eq!(out, vec!["Datagram Message"]);
}

#[test]
fn test_php_stream_socket_pair_close_one_side_signals_eof() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
fclose($pair[0]);
$data = fread($pair[1], 10);
$eof = feof($pair[1]);
fclose($pair[1]);
echo $data === "" && $eof ? "CLOSED_SIDE_EOF_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_pair_large_payload_transfer() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
$payload = str_repeat("ABCDEFGH", 1024); // 8KB payload
fwrite($pair[0], $payload);
$received = stream_get_contents($pair[1], strlen($payload));
fclose($pair[0]);
fclose($pair[1]);
echo $received === $payload ? "8KB_PAIR_TRANSFER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_pair_select_read_readiness() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
fwrite($pair[0], "ready_data");

$read = [$pair[1]];
$write = null;
$except = null;
$changed = stream_select($read, $write, $except, 1);

fclose($pair[0]);
fclose($pair[1]);
echo $changed === 1 && count($read) === 1 ? "STREAM_SELECT_READ_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_pair_unclosed_buffer_flush() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
fwrite($pair[0], "line1\n");
fwrite($pair[0], "line2\n");
$l1 = fgets($pair[1]);
$l2 = fgets($pair[1]);
fclose($pair[0]);
fclose($pair[1]);
echo trim($l1) === "line1" && trim($l2) === "line2" ? "FGETS_PAIR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_pair_invalid_domain_returns_false() {
    compile_ok(
        r##"<?php
$res = @stream_socket_pair(99999, STREAM_SOCK_STREAM, 0);
echo $res === false ? "INVALID_DOMAIN_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_pair_stream_copy_to_stream() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
$mem = fopen("php://memory", "r+");
fwrite($mem, "Stream Copy Test Data");
rewind($mem);

stream_copy_to_stream($mem, $pair[0]);
$copied = stream_get_contents($pair[1], 21);

fclose($mem);
fclose($pair[0]);
fclose($pair[1]);
echo $copied === "Stream Copy Test Data" ? "STREAM_COPY_PAIR_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_pair_metadata_type() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
$meta = stream_get_meta_data($pair[0]);
fclose($pair[0]);
fclose($pair[1]);
echo isset($meta["stream_type"]) ? "PAIR_META_TYPE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_stream_socket_pair_partial_read_offset() {
    compile_ok(
        r##"<?php
$pair = stream_socket_pair(STREAM_PF_UNIX, STREAM_SOCK_STREAM, 0);
fwrite($pair[0], "0123456789");
$part1 = fread($pair[1], 4);
$part2 = fread($pair[1], 6);
fclose($pair[0]);
fclose($pair[1]);
echo $part1 === "0123" && $part2 === "456789" ? "PARTIAL_READ_PAIR_OK" : "FAIL";
"##,
    );
}

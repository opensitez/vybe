use super::helpers::run_prints;

crate::php_cases! {
    stream_socket_accept_timeout => {
        r#"<?php
$server = stream_socket_server("tcp://127.0.0.1:0", $errno, $errstr);
$start = microtime(true);
$conn = @stream_socket_accept($server, 0.1);
$end = microtime(true);

echo is_resource($server) ? "server|" : "no|";
echo is_resource($conn) ? "conn" : "timeout";
"#,
        ["server|timeout"]
    };
}

use super::helpers::run_prints;

crate::php_cases! {
    stream_select_read_timeout => {
        r#"<?php
$server = stream_socket_server("tcp://127.0.0.1:0");
$read = [$server];
$write = null;
$except = null;
$num = @stream_select($read, $write, $except, 0, 100000);
echo $num;
"#,
        ["0"]
    };

    stream_select_write_ready => {
        r#"<?php
$fp = fopen("php://temp", "w+");
$read = null;
$write = [$fp];
$except = null;
$num = stream_select($read, $write, $except, 0);
echo $num;
"#,
        ["1"]
    };
}

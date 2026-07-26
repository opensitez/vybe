
crate::php_cases! {
    stream_set_chunk_size => {
        r#"<?php
$fp = fopen("php://temp", "w+");
$res = stream_set_chunk_size($fp, 1024);
echo $res;
"#,
        ["1024"]
    };
}

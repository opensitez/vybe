
crate::php_cases! {
    stream_copy_to_stream_offset => {
        r#"<?php
$src = fopen("php://memory", "w+");
fwrite($src, "Hello, world!");

$dest = fopen("php://memory", "w+");
stream_copy_to_stream($src, $dest, 5, 7);

rewind($dest);
echo stream_get_contents($dest);
"#,
        ["world"]
    };
}

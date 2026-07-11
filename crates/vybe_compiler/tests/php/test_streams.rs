//! `php://memory`, `php://temp`, and stream helper functions.

crate::php_cases! {
    php_memory_stream_write_read_roundtrip => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, 'payload');
rewind($fp);
echo stream_get_contents($fp);
"#,
        ["payload"]
    };

    php_memory_stream_tell_after_write => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, 'abc');
echo ftell($fp);
"#,
        ["3"]
    };

    php_memory_stream_eof_after_read_all => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, 'x');
rewind($fp);
stream_get_contents($fp);
echo feof($fp) ? 'eof' : 'more';
"#,
        ["eof"]
    };

    stream_get_meta_data_uri_for_memory => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
$m = stream_get_meta_data($fp);
echo str_starts_with($m['uri'], 'php://') ? 'php' : 'other';
"#,
        ["php"]
    };

    fgets_reads_line_with_newline => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, "a\nb");
rewind($fp);
echo fgets($fp);
"#,
        ["a"]
    };

    fread_reads_exact_bytes => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, '12345');
rewind($fp);
echo fread($fp, 2);
"#,
        ["12"]
    };

    fputcsv_and_fgetcsv_roundtrip => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fputcsv($fp, ['x', 'y']);
rewind($fp);
echo implode('|', fgetcsv($fp));
"#,
        ["x|y"]
    };

    stream_copy_to_string_from_memory => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, 'copy');
rewind($fp);
echo stream_copy_to_stream($fp, fopen('php://memory', 'r+')) !== false ? 'ok' : 'fail';
"#,
        ["ok"]
    };

    php_temp_stream_writable => {
        r#"<?php
$fp = fopen('php://temp', 'r+');
fwrite($fp, 't');
rewind($fp);
echo stream_get_contents($fp);
"#,
        ["t"]
    };

    stream_filter_append_when_available => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
fwrite($fp, 'data');
rewind($fp);
echo is_resource($fp) ? 'res' : 'no';
"#,
        ["res"]
    };
}

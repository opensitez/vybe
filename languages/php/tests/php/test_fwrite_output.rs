//! `fwrite` and `php://output` — byte writes to stdout stream handles.

crate::php_cases! {
    fwrite_stdout_concatenates_with_echo => {
        r#"<?php
fwrite(STDOUT, 'fw');
echo 'out';
"#,
        ["fwout"]
    };

    fwrite_stdout_two_writes_no_newline => {
        r#"<?php
fwrite(STDOUT, 'ab');
fwrite(STDOUT, 'cd');
"#,
        ["abcd"]
    };

    fwrite_stdout_returns_byte_count => {
        r#"<?php
$n = fwrite(STDOUT, 'four');
echo ':';
echo $n;
"#,
        ["four:4"]
    };

    fwrite_php_output_stream_concatenates => {
        r#"<?php
$fp = fopen('php://output', 'w');
fwrite($fp, 'stream');
fclose($fp);
echo '!';
"#,
        ["stream!"]
    };

    fwrite_php_output_with_explicit_newline => {
        r#"<?php
$fp = fopen('php://output', 'w');
fwrite($fp, "row\n");
fclose($fp);
echo 'end';
"#,
        ["row", "end"]
    };
}

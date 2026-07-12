//! `fprintf` / `vfprintf` — formatted writes to PHP stream handles (`STDOUT`,
//! `STDERR`, file pointers). Runtime tests assert PHP-spec stdout bytes; failures
//! mark missing profile/emitter/stream wiring.

crate::php_cases! {
    fprintf_stdout_string_concatenates_with_echo => {
        r#"<?php
fprintf(STDOUT, '%s', 'fp');
echo 'out';
"#,
        ["fpout"]
    };

    fprintf_stdout_integer_format => {
        r#"<?php
fprintf(STDOUT, '%d', 42);
"#,
        ["42"]
    };

    fprintf_stdout_string_and_integer => {
        r#"<?php
fprintf(STDOUT, '%s=%d', 'count', 9);
"#,
        ["count=9"]
    };

    fprintf_stdout_float_precision => {
        r#"<?php
fprintf(STDOUT, '%.2f', 3.14159);
"#,
        ["3.14"]
    };

    fprintf_stdout_two_calls_no_implicit_newline => {
        r#"<?php
fprintf(STDOUT, 'a');
fprintf(STDOUT, 'b');
"#,
        ["ab"]
    };

    fprintf_stdout_embedded_newline_then_echo => {
        r#"<?php
fprintf(STDOUT, "line\n");
echo 'next';
"#,
        ["line", "next"]
    };

    fprintf_stdout_percent_literal => {
        r#"<?php
fprintf(STDOUT, '%%');
echo 'ok';
"#,
        ["%ok"]
    };

    fprintf_stdout_hex_lowercase => {
        r#"<?php
fprintf(STDOUT, '%x', 255);
"#,
        ["ff"]
    };

    fprintf_stdout_zero_padded_width => {
        r#"<?php
fprintf(STDOUT, '%05d', 7);
"#,
        ["00007"]
    };

    fprintf_stdout_returns_byte_count_on_stdout => {
        r#"<?php
$written = fprintf(STDOUT, '%s', 'xy');
echo ':';
echo $written;
"#,
        ["xy:2"]
    };

    fprintf_stdout_mixed_with_printf_same_line => {
        r#"<?php
fprintf(STDOUT, 'F');
printf('P');
echo '!';
"#,
        ["FP!"]
    };

    fprintf_stdout_item_price_format => {
        r#"<?php
fprintf(STDOUT, 'Item: %s costs $%.2f', 'widget', 4.99);
"#,
        ["Item: widget costs $4.99"]
    };

    vfprintf_stdout_string_and_integer_array => {
        r#"<?php
vfprintf(STDOUT, '%s-%d', ['vybe', 2]);
echo '!';
"#,
        ["vybe-2!"]
    };

    vfprintf_stdout_three_placeholders => {
        r#"<?php
vfprintf(STDOUT, '%s %d %s', ['a', 1, 'b']);
"#,
        ["a 1 b"]
    };

    fprintf_stderr_does_not_pollute_stdout_capture => {
        r#"<?php
fprintf(STDERR, 'on-stderr');
echo 'on-stdout';
"#,
        ["on-stdout"]
    };

    fprintf_stdout_sign_flag_positive => {
        r#"<?php
fprintf(STDOUT, '%+d', 42);
"#,
        ["+42"]
    };

    fprintf_stdout_left_aligned_string => {
        r#"<?php
fprintf(STDOUT, '%-5s|', 'hi');
"#,
        ["hi   |"]
    };
}

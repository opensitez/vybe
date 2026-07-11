//! Runtime output semantics for `echo`, `print`, `printf`, `sprintf`, and related
//! primitives that write directly to stdout. See dedicated files for other
//! surfaces: `test_fprintf.rs`, `test_fwrite_output.rs`, `test_var_dump_output.rs`,
//! `test_print_r_output.rs`, `test_var_export_output.rs`, `test_error_log_output.rs`.

crate::php_cases! {
    // ── echo ─────────────────────────────────────────────────────────────

    echo_two_calls_concatenate_without_newline => {
        r#"<?php
echo 'yes';
echo 'no';
"#,
        ["yesno"]
    };

    echo_comma_syntax_concatenates_args => {
        r#"<?php
echo 'hello', ' ', 'world';
"#,
        ["hello world"]
    };

    echo_comma_multiple_types => {
        r#"<?php
echo 'n=', 42, ' b=', true;
"#,
        ["n=42 b=1"]
    };

    echo_explicit_lf_splits_lines => {
        r#"<?php
echo "line1\n";
echo 'line2';
"#,
        ["line1", "line2"]
    };

    echo_no_implicit_carriage_return => {
        r#"<?php
echo 'x';
echo 'y';
"#,
        ["xy"]
    };

    echo_literal_cr_without_lf_one_line => {
        r#"<?php
echo "a\rb";
"#,
        ["a", "b"]
    };

    echo_crlf_splits_lines => {
        r#"<?php
echo "first\r\n";
echo 'second';
"#,
        ["first", "second"]
    };

    echo_false_null_true_coercion => {
        r#"<?php
echo false;
echo null;
echo true;
"#,
        ["1"]
    };

    echo_interpolated_double_quoted => {
        r#"<?php
$name = 'vybe';
echo "hi $name";
echo '!';
"#,
        ["hi vybe!"]
    };

    // ── print (statement and function) ───────────────────────────────────

    print_statement_two_calls_concatenate => {
        r#"<?php
print 'foo';
print 'bar';
"#,
        ["foobar"]
    };

    print_function_two_calls_concatenate => {
        r#"<?php
print('foo');
print('bar');
"#,
        ["foobar"]
    };

    print_return_one_on_stdout_after_payload => {
        r#"<?php
$n = print 'x';
echo $n;
"#,
        ["x1"]
    };

    print_then_echo_no_separator => {
        r#"<?php
print 'hi';
echo '!';
"#,
        ["hi!"]
    };

    echo_then_print_no_separator => {
        r#"<?php
echo 'go';
print 'ing';
"#,
        ["going"]
    };

    print_explicit_lf_splits_lines => {
        r#"<?php
print "one\n";
print 'two';
"#,
        ["one", "two"]
    };

    // ── printf family ────────────────────────────────────────────────────

    printf_without_newline_then_echo => {
        r#"<?php
printf('%s', 'fmt');
echo 'tail';
"#,
        ["fmttail"]
    };

    printf_embedded_lf_splits_lines => {
        r#"<?php
printf("a\n");
echo 'b';
"#,
        ["a", "b"]
    };

    vprintf_writes_formatted_without_extra_newline => {
        r#"<?php
vprintf('%s-%d', ['vybe', 2]);
echo '!';
"#,
        ["vybe-2!"]
    };

    // ── sprintf family (return string, no stdout) ──────────────────────────

    sprintf_does_not_write_stdout => {
        r#"<?php
$s = sprintf('only-%d', 9);
echo $s;
"#,
        ["only-9"]
    };

    echo_sprintf_results_concatenate => {
        r#"<?php
echo sprintf('%d', 7);
echo sprintf('%d', 3);
"#,
        ["73"]
    };

    vsprintf_returns_string_without_stdout => {
        r#"<?php
$s = vsprintf('%s:%d', ['x', 5]);
echo $s;
"#,
        ["x:5"]
    };

    // ── var_export return-to-string ──────────────────────────────────────

    var_export_return_true_scalar => {
        r#"<?php
echo var_export(42, true);
"#,
        ["42"]
    };

    var_export_return_true_bool_false => {
        r#"<?php
echo var_export(false, true);
"#,
        ["false"]
    };

    // ── json_encode (string to stdout via echo) ──────────────────────────

    json_encode_to_stdout_without_extra_newline => {
        r#"<?php
echo json_encode(['z' => 9]);
echo '!';
"#,
        ["{\"z\":9}!"]
    };

    // ── mixed primitives ─────────────────────────────────────────────────

    echo_print_printf_chain_one_line => {
        r#"<?php
echo '1';
print '2';
printf('%d', 3);
echo '4';
"#,
        ["1234"]
    };
}

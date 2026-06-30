//! Error reporting, logging, and `error_*` helpers (non-fatal paths).

crate::php_cases! {
    error_reporting_get_returns_int => {
        r#"<?php
echo is_int(error_reporting()) ? 'int' : 'no';
"#,
        ["int"]
    };

    error_reporting_set_and_restore => {
        r#"<?php
$old = error_reporting(E_ALL);
error_reporting($old);
echo error_reporting() === $old ? 'same' : 'diff';
"#,
        ["same"]
    };

    trigger_error_user_notice => {
        r#"<?php
set_error_handler(fn() => true);
trigger_error('msg', E_USER_NOTICE);
restore_error_handler();
echo 'ok';
"#,
        ["ok"]
    };

    trigger_error_user_warning => {
        r#"<?php
set_error_handler(fn() => true);
trigger_error('warn', E_USER_WARNING);
restore_error_handler();
echo 'ok';
"#,
        ["ok"]
    };

    set_error_handler_returns_previous => {
        r#"<?php
$h = fn() => true;
$prev = set_error_handler($h);
restore_error_handler();
echo $prev === null ? 'null' : 'fn';
"#,
        ["null"]
    };

    restore_error_handler_after_set => {
        r#"<?php
set_error_handler(fn() => true);
restore_error_handler();
echo 'restored';
"#,
        ["restored"]
    };

    set_exception_handler_callable => {
        r#"<?php
set_exception_handler(fn($e) => null);
restore_exception_handler();
echo 'ok';
"#,
        ["ok"]
    };

    error_get_last_after_trigger => {
        r#"<?php
set_error_handler(fn() => true);
trigger_error('x', E_USER_NOTICE);
restore_error_handler();
$e = error_get_last();
echo isset($e['message']) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    error_clear_last => {
        r#"<?php
set_error_handler(fn() => true);
trigger_error('x', E_USER_NOTICE);
restore_error_handler();
error_clear_last();
echo error_get_last() === null ? 'cleared' : 'set';
"#,
        ["cleared"]
    };

    error_log_returns_true => {
        r#"<?php
echo error_log('test', 3, sys_get_temp_dir() . '/vybe_php_test.log') ? '1' : '0';
"#,
        ["1"]
    };

    debug_backtrace_has_file_key => {
        r#"<?php
$bt = debug_backtrace();
echo isset($bt[0]['file']) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    debug_print_backtrace_captures => {
        r#"<?php
ob_start();
debug_print_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS, 1);
$out = ob_get_clean();
echo strlen($out) > 0 ? 'trace' : 'empty';
"#,
        ["trace"]
    };

    get_debug_type_string => {
        r#"<?php
echo get_debug_type('s');
"#,
        ["string"]
    };

    get_debug_type_int => {
        r#"<?php
echo get_debug_type(1);
"#,
        ["int"]
    };

    get_debug_type_object => {
        r#"<?php
echo get_debug_type(new stdClass());
"#,
        ["stdClass"]
    };

    get_debug_type_null => {
        r#"<?php
echo get_debug_type(null);
"#,
        ["null"]
    };

    get_resource_type_stream => {
        r#"<?php
$f = fopen('php://memory', 'r+');
echo get_resource_type($f);
"#,
        ["stream"]
    };

    is_int_checks_integer => {
        r#"<?php
echo is_int(42) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_float_checks_double => {
        r#"<?php
echo is_float(1.5) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_bool_checks_boolean => {
        r#"<?php
echo is_bool(false) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_null_checks_null => {
        r#"<?php
echo is_null(null) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_array_checks_array => {
        r#"<?php
echo is_array([]) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_object_checks_object => {
        r#"<?php
echo is_object(new stdClass()) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_string_checks_string => {
        r#"<?php
echo is_string('x') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_callable_on_closure => {
        r#"<?php
echo is_callable(fn() => 1) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_iterable_on_array => {
        r#"<?php
echo is_iterable([1]) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_countable_on_array => {
        r#"<?php
echo is_countable([]) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    gettype_integer => {
        r#"<?php
echo gettype(0);
"#,
        ["integer"]
    };

    gettype_double => {
        r#"<?php
echo gettype(0.0);
"#,
        ["double"]
    };

    var_export_string_roundtrip => {
        r#"<?php
echo var_export(5, true);
"#,
        ["5"]
    };

    print_r_return_string => {
        r#"<?php
$out = print_r(['a' => 1], true);
echo str_contains($out, '[a]') ? 'has' : 'no';
"#,
        ["has"]
    };

    highlight_string_wraps_php => {
        r#"<?php
ob_start();
highlight_string('<?php echo 1;');
$out = ob_get_clean();
echo str_contains($out, 'php') ? 'html' : 'no';
"#,
        ["html"]
    };

    ini_get_display_errors => {
        r#"<?php
echo is_string(ini_get('display_errors')) ? 'str' : 'no';
"#,
        ["str"]
    };

    ini_set_display_errors => {
        r#"<?php
$old = ini_set('display_errors', '0');
ini_set('display_errors', $old ?: '1');
echo 'ok';
"#,
        ["ok"]
    };

    assert_options_get => {
        r#"<?php
echo is_int(assert_options(ASSERT_ACTIVE)) ? 'int' : 'no';
"#,
        ["int"]
    };

    user_error_handler_receives_errno => {
        r#"<?php
$code = 0;
set_error_handler(function ($errno) use (&$code) { $code = $errno; return true; });
trigger_error('t', E_USER_WARNING);
restore_error_handler();
echo $code === E_USER_WARNING ? 'match' : 'miss';
"#,
        ["match"]
    };
}

//! `is_*` type checks, `gettype`, and related introspection.

crate::php_cases! {
    is_array_true_for_list => {
        r#"<?php
echo is_array([1, 2]) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_array_false_for_string => {
        r#"<?php
echo is_array('x') ? 'yes' : 'no';
"#,
        ["no"]
    };

    is_int_detects_integer => {
        r#"<?php
echo is_int(7) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_float_detects_double => {
        r#"<?php
echo is_float(1.5) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_string_detects_text => {
        r#"<?php
echo is_string('a') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_bool_true_and_false => {
        r#"<?php
echo (is_bool(true) ? 't' : 'f') . (is_bool(1) ? 't' : 'f');
"#,
        ["tf"]
    };

    is_null_only_for_null => {
        r#"<?php
echo is_null(null) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_object_for_stdclass => {
        r#"<?php
echo is_object(new stdClass()) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_callable_for_function_name_string => {
        r#"<?php
echo is_callable('strlen') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_callable_for_closure => {
        r#"<?php
echo is_callable(fn() => 1) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_numeric_accepts_numeric_string => {
        r#"<?php
echo is_numeric('42') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_numeric_rejects_alpha => {
        r#"<?php
echo is_numeric('42a') ? 'yes' : 'no';
"#,
        ["no"]
    };

    is_scalar_includes_string_not_array => {
        r#"<?php
echo (is_scalar('x') ? 's' : '-') . (is_scalar([]) ? 'a' : 'n');
"#,
        ["sn"]
    };

    is_countable_array_yes_string_no => {
        r#"<?php
echo (is_countable([]) ? 'a' : '-') . (is_countable('x') ? 's' : 'n');
"#,
        ["an"]
    };

    is_iterable_array_yes_int_no => {
        r#"<?php
echo (is_iterable([1]) ? 'a' : '-') . (is_iterable(1) ? 'i' : 'n');
"#,
        ["an"]
    };

    gettype_integer_label => {
        r#"<?php
echo gettype(1);
"#,
        ["integer"]
    };

    gettype_double_label => {
        r#"<?php
echo gettype(1.0);
"#,
        ["double"]
    };

    gettype_string_label => {
        r#"<?php
echo gettype('x');
"#,
        ["string"]
    };

    gettype_array_label => {
        r#"<?php
echo gettype([]);
"#,
        ["array"]
    };

    gettype_object_label => {
        r#"<?php
echo gettype(new stdClass());
"#,
        ["object"]
    };

    gettype_null_label => {
        r#"<?php
echo gettype(null);
"#,
        ["NULL"]
    };

    is_finite_for_normal_float => {
        r#"<?php
echo is_finite(1.5) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_infinite_detects_inf => {
        r#"<?php
echo is_infinite(INF) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_nan_detects_nan => {
        r#"<?php
echo is_nan(NAN) ? 'yes' : 'no';
"#,
        ["yes"]
    };
}

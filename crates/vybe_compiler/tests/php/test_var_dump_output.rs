//! `var_dump` stdout formatting and concatenation (PHP spec).

crate::php_cases! {
    var_dump_string_type_and_length_format => {
        r#"<?php
var_dump('x');
echo '|';
"#,
        ["string(1) \"x\"|"]
    };

    var_dump_integer_format => {
        r#"<?php
var_dump(42);
echo '|';
"#,
        ["int(42)|"]
    };

    var_dump_bool_true_format => {
        r#"<?php
var_dump(true);
echo '|';
"#,
        ["bool(true)|"]
    };

    var_dump_null_format => {
        r#"<?php
var_dump(null);
echo '|';
"#,
        ["NULL|"]
    };

    var_dump_two_scalars_concatenate_on_one_line => {
        r#"<?php
var_dump('a');
var_dump('b');
"#,
        ["string(1) \"a\"string(1) \"b\""]
    };
}

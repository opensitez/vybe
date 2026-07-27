crate::php_cases! {
    filter_var_boolean_true_values => {
        r#"<?php
echo filter_var("yes", FILTER_VALIDATE_BOOLEAN, FILTER_NULL_ON_FAILURE) === true ? "true" : "fail";
echo "|";
echo filter_var("on", FILTER_VALIDATE_BOOLEAN, FILTER_NULL_ON_FAILURE) === true ? "true" : "fail";
"#,
        ["true|true"]
    };

    filter_var_boolean_null_on_failure => {
        r#"<?php
echo filter_var("maybe", FILTER_VALIDATE_BOOLEAN, FILTER_NULL_ON_FAILURE) === null ? "null" : "fail";
"#,
        ["null"]
    };
}

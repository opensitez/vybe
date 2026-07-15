use super::helpers::run_prints;

crate::php_cases! {
    datetime_create_from_format_basic => {
        r#"<?php
$dt = DateTime::createFromFormat('j-M-Y', '15-Feb-2009');
echo $dt->format('Y-m-d');
"#,
        ["2009-02-15"]
    };

    datetime_create_from_format_strict_pipe => {
        r#"<?php
// The '|' resets all fields to the Unix Epoch
$dt = DateTime::createFromFormat('Y-m-d|', '2009-02-15');
echo $dt->format('Y-m-d H:i:s');
"#,
        ["2009-02-15 00:00:00"]
    };
}


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

    datetime_create_from_format_timezone => {
        r#"<?php
$dt = DateTime::createFromFormat('Y-m-d H:i O', '2024-11-02 14:30 +0200');
echo $dt->format('Y-m-d H:i:sP');
"#,
        ["2024-11-02 14:30:00+02:00"]
    };

    datetime_create_from_format_immutable => {
        r#"<?php
$dt = DateTimeImmutable::createFromFormat('d/m/Y H:i', '31/12/2024 23:59');
echo $dt->format('Y-m-d H:i');
"#,
        ["2024-12-31 23:59"]
    };

    datetime_create_from_format_epoch => {
        r#"<?php
$dt = DateTimeImmutable::createFromFormat('U', '0');
echo $dt->format('Y-m-d H:i:s');
"#,
        ["1970-01-01 00:00:00"]
    };

    datetime_create_from_format_strict_invalid => {
        r#"<?php
$dt = DateTime::createFromFormat('Y-m-d', '2024-99-99');
$errors = DateTime::getLastErrors();
echo $dt === false ? 'false' : 'not false';
echo ':';
echo $errors['error_count'] > 0 ? 'has_errors' : 'no_errors';
"#,
        ["false:has_errors"]
    };

    datetime_create_from_format_exclamation_reset => {
        r#"<?php
$base = new DateTime('1999-12-31 11:22:33');
$dt = DateTime::createFromFormat('!Y-m-d', '2000-01-02', $base->getTimezone());
echo $dt->format('Y-m-d H:i:s');
"#,
        ["2000-01-02 00:00:00"]
    };
}

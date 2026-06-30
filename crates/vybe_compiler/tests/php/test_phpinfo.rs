//! `phpinfo()` and related runtime introspection output.

crate::php_cases! {
    phpversion_matches_php_version_constant => {
        r#"<?php
echo phpversion() === PHP_VERSION ? 'match' : 'diff';
"#,
        ["match"]
    };

    php_version_id_is_80000 => {
        r#"<?php
echo PHP_VERSION_ID === 80000 ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    php_sapi_name_is_cli_in_tests => {
        r#"<?php
echo php_sapi_name();
"#,
        ["cli"]
    };

    php_os_constant_is_darwin => {
        r#"<?php
echo PHP_OS;
"#,
        ["Darwin"]
    };

    php_int_size_is_eight => {
        r#"<?php
echo PHP_INT_SIZE;
"#,
        ["8"]
    };

    phpinfo_emits_version_system_and_sapi_lines => {
        r#"<?php
phpinfo();
echo 'END';
"#,
        [
            "phpinfo()",
            "PHP Version => 8.0.0",
            "System => Darwin",
            "Build Date => vybe",
            "Server API => cli",
            "PHP API => vybex",
            "PHP Extension Build => vybe",
            "Zend Extension Build => n/a",
            "PHP Integer Size => 8",
            "END",
        ]
    };

    phpinfo_return_value_is_true => {
        r#"<?php
$ok = phpinfo();
echo $ok ? 'true' : 'false';
"#,
        [
            "phpinfo()",
            "PHP Version => 8.0.0",
            "System => Darwin",
            "Build Date => vybe",
            "Server API => cli",
            "PHP API => vybex",
            "PHP Extension Build => vybe",
            "Zend Extension Build => n/a",
            "PHP Integer Size => 8",
            "true",
        ]
    };

    phpinfo_flags_argument_accepted => {
        r#"<?php
phpinfo(INFO_GENERAL);
echo 'done';
"#,
        [
            "phpinfo()",
            "PHP Version => 8.0.0",
            "System => Darwin",
            "Build Date => vybe",
            "Server API => cli",
            "PHP API => vybex",
            "PHP Extension Build => vybe",
            "Zend Extension Build => n/a",
            "PHP Integer Size => 8",
            "done",
        ]
    };

    extension_loaded_mysqli_reports_status => {
        r#"<?php
echo extension_loaded('mysqli') ? 'mysqli-yes' : 'mysqli-no';
"#,
        ["mysqli-yes"]
    };
}

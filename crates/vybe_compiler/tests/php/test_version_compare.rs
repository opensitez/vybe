//! `version_compare`, `phpversion`, and `extension_loaded`.

crate::php_cases! {
    version_compare_greater_than => {
        r#"<?php
echo version_compare('8.2.0', '8.1.0', '>') ? 'gt' : 'no';
"#,
        ["gt"]
    };

    version_compare_equal_normalized => {
        r#"<?php
echo version_compare('8.2', '8.2.0', '==') ? 'eq' : 'ne';
"#,
        ["eq"]
    };

    version_compare_less_or_equal => {
        r#"<?php
echo version_compare('7.4.33', '8.0.0', '<=') ? 'le' : 'no';
"#,
        ["le"]
    };

    version_compare_not_equal => {
        r#"<?php
echo version_compare('1.0.0', '1.0.1', '!=') ? 'ne' : 'eq';
"#,
        ["ne"]
    };

    phpversion_returns_non_empty_string => {
        r#"<?php
echo strlen(PHP_VERSION) > 0 ? 'set' : 'empty';
"#,
        ["set"]
    };

    phpversion_matches_php_version_constant => {
        r#"<?php
echo phpversion() === PHP_VERSION ? 'match' : 'diff';
"#,
        ["match"]
    };

    extension_loaded_detects_core_extension => {
        r#"<?php
echo extension_loaded('standard') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    extension_loaded_false_for_missing => {
        r#"<?php
echo extension_loaded('definitely_not_a_real_extension') ? 'yes' : 'no';
"#,
        ["no"]
    };

    get_loaded_extensions_includes_standard => {
        r#"<?php
echo in_array('standard', get_loaded_extensions(), true) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    version_compare_pre_release_lower_than_release => {
        r#"<?php
echo version_compare('8.2.0beta1', '8.2.0', '<') ? 'beta' : 'rel';
"#,
        ["beta"]
    };

    php_sapi_name_non_empty => {
        r#"<?php
echo strlen(php_sapi_name()) > 0 ? 'sapi' : 'empty';
"#,
        ["sapi"]
    };

    php_uname_returns_string => {
        r#"<?php
echo is_string(php_uname()) ? 'str' : 'other';
"#,
        ["str"]
    };

    zend_version_constant_non_empty => {
        r#"<?php
echo strlen(ZEND_VERSION) > 0 ? 'zend' : 'empty';
"#,
        ["zend"]
    };
}

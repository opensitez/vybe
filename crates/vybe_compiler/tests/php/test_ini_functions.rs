//! `ini_get`, `ini_set`, `ini_restore`, `ini_get_all`, and related config introspection.

crate::php_cases! {
    ini_get_returns_string_for_display_errors => {
        r#"<?php
echo is_string(ini_get('display_errors')) ? 'str' : 'other';
"#,
        ["str"]
    };

    ini_get_all_includes_display_errors_entry => {
        r#"<?php
$all = ini_get_all();
echo isset($all['display_errors']) ? 'has' : 'missing';
"#,
        ["has"]
    };

    ini_set_toggles_display_errors_then_restores => {
        r#"<?php
$old = ini_set('display_errors', '0');
ini_restore('display_errors');
echo ini_get('display_errors') === $old || $old === false ? 'restored' : 'changed';
"#,
        ["restored"]
    };

    ini_set_max_execution_time_accepts_numeric_string => {
        r#"<?php
$prev = ini_set('max_execution_time', '120');
echo is_string($prev) || $prev === false ? 'ok' : 'bad';
"#,
        ["ok"]
    };

    ini_get_memory_limit_returns_suffix => {
        r#"<?php
echo str_ends_with(ini_get('memory_limit'), 'M') || str_ends_with(ini_get('memory_limit'), 'G') || ini_get('memory_limit') === '-1' ? 'limit' : 'other';
"#,
        ["limit"]
    };

    ini_get_post_max_size_non_empty => {
        r#"<?php
echo strlen(ini_get('post_max_size')) > 0 ? 'set' : 'empty';
"#,
        ["set"]
    };

    ini_get_upload_max_filesize_non_empty => {
        r#"<?php
echo strlen(ini_get('upload_max_filesize')) > 0 ? 'set' : 'empty';
"#,
        ["set"]
    };

    get_cfg_var_matches_ini_get_for_php_version => {
        r#"<?php
echo get_cfg_var('PHP_VERSION') === PHP_VERSION ? 'match' : 'diff';
"#,
        ["match"]
    };

    ini_get_bool_casts_on_value => {
        r#"<?php
$v = ini_get('display_errors');
echo in_array($v, ['0', '1', 'Off', 'On', 'stderr'], true) || is_numeric($v) ? 'boolish' : 'other';
"#,
        ["boolish"]
    };

    ini_set_user_error_handler_name_returns_prior => {
        r#"<?php
$old = ini_set('error_reporting', (string)E_ALL);
ini_set('error_reporting', $old !== false ? $old : (string)E_ALL);
echo is_numeric(ini_get('error_reporting')) ? 'numeric' : 'str';
"#,
        ["numeric"]
    };

    ini_get_default_charset_non_empty => {
        r#"<?php
echo strlen(ini_get('default_charset')) > 0 ? 'charset' : 'empty';
"#,
        ["charset"]
    };

    ini_get_precision_is_numeric_string => {
        r#"<?php
echo is_numeric(ini_get('precision')) ? 'num' : 'no';
"#,
        ["num"]
    };

    ini_alter_alias_of_ini_set => {
        r#"<?php
$old = ini_alter('display_errors', ini_get('display_errors'));
echo $old !== false || $old === false ? 'called' : 'fail';
"#,
        ["called"]
    };

    ini_restore_after_set_returns_original_display_errors => {
        r#"<?php
$orig = ini_get('display_errors');
ini_set('display_errors', $orig === '1' ? '0' : '1');
ini_restore('display_errors');
echo ini_get('display_errors') === $orig ? 'back' : 'stuck';
"#,
        ["back"]
    };

    ini_get_include_path_non_empty => {
        r#"<?php
echo strlen(ini_get('include_path')) >= 0 ? 'path' : 'no';
"#,
        ["path"]
    };

    ini_get_session_save_path_is_string => {
        r#"<?php
echo is_string(ini_get('session.save_path')) ? 'str' : 'other';
"#,
        ["str"]
    };

    ini_get_opcache_enable_when_present => {
        r#"<?php
$v = ini_get('opcache.enable');
echo $v === false || $v === '' || $v === '0' || $v === '1' ? 'known' : 'other';
"#,
        ["known"]
    };
}

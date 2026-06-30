//! Session API: `session_start`, `$_SESSION`, handlers (runtime shape).

crate::php_cases! {
    session_status_not_active_before_start => {
        r#"<?php
echo session_status() === PHP_SESSION_NONE ? 'none' : 'other';
"#,
        ["none"]
    };

    session_start_activates_session => {
        r#"<?php
session_start();
echo session_status() === PHP_SESSION_ACTIVE ? 'active' : 'no';
"#,
        ["active"]
    };

    session_id_returns_string => {
        r#"<?php
session_start();
$id = session_id();
echo is_string($id) ? 'str' : 'no';
"#,
        ["str"]
    };

    session_name_default_is_php => {
        r#"<?php
echo session_name();
"#,
        ["PHPSESSID"]
    };

    session_name_custom_before_start => {
        r#"<?php
session_name('APPSESSID');
session_start();
echo session_name();
"#,
        ["APPSESSID"]
    };

    session_set_get_cookie_params_lifetime => {
        r#"<?php
session_set_cookie_params(['lifetime' => 3600]);
session_start();
$p = session_get_cookie_params();
echo $p['lifetime'];
"#,
        ["3600"]
    };

    session_write_read_roundtrip => {
        r#"<?php
session_start();
$_SESSION['k'] = 'v';
session_write_close();
session_start();
echo $_SESSION['k'] ?? 'missing';
"#,
        ["v"]
    };

    session_unset_clears_superglobal => {
        r#"<?php
session_start();
$_SESSION['a'] = 1;
session_unset();
echo empty($_SESSION) ? 'empty' : 'set';
"#,
        ["empty"]
    };

    session_destroy_ends_session => {
        r#"<?php
session_start();
$_SESSION['x'] = 1;
session_destroy();
echo session_status() === PHP_SESSION_NONE ? 'gone' : 'live';
"#,
        ["gone"]
    };

    session_regenerate_id_changes_id => {
        r#"<?php
session_start();
$old = session_id();
session_regenerate_id(true);
echo $old !== session_id() ? 'new' : 'same';
"#,
        ["new"]
    };

    session_cache_limiter_nocache => {
        r#"<?php
session_cache_limiter('nocache');
echo session_cache_limiter();
"#,
        ["nocache"]
    };

    session_cache_expire_minutes => {
        r#"<?php
session_cache_expire(30);
echo session_cache_expire();
"#,
        ["30"]
    };

    session_module_name_files => {
        r#"<?php
echo session_module_name();
"#,
        ["files"]
    };

    session_save_path_get_set => {
        r#"<?php
$path = sys_get_temp_dir();
session_save_path($path);
echo session_save_path() === $path ? 'ok' : 'no';
"#,
        ["ok"]
    };

    session_encode_decode_roundtrip => {
        r#"<?php
session_start();
$_SESSION['n'] = 42;
$blob = session_encode();
session_unset();
session_decode($blob);
echo $_SESSION['n'] ?? 0;
"#,
        ["42"]
    };

    session_abort_discards_changes => {
        r#"<?php
session_start();
$_SESSION['tmp'] = 1;
session_abort();
session_start();
echo isset($_SESSION['tmp']) ? 'yes' : 'no';
"#,
        ["no"]
    };

    session_reset_restores_snapshot => {
        r#"<?php
session_start();
$_SESSION['a'] = 1;
session_reset();
echo isset($_SESSION['a']) ? 'yes' : 'no';
"#,
        ["no"]
    };

    session_gc_probability_returns_int => {
        r#"<?php
echo is_int(session_gc(['probability' => 1, 'divisor' => 100])) ? 'int' : 'no';
"#,
        ["int"]
    };

    session_create_id_unique_prefix => {
        r#"<?php
$id = session_create_id('app');
echo str_starts_with($id, 'app') ? 'pref' : 'nop';
"#,
        ["pref"]
    };

    session_set_save_handler_user_array => {
        r#"<?php
session_set_save_handler([
    'open' => fn() => true,
    'close' => fn() => true,
    'read' => fn($id) => '',
    'write' => fn($id, $data) => true,
    'destroy' => fn($id) => true,
    'gc' => fn($max) => 0,
]);
session_start();
echo 'handler';
"#,
        ["handler"]
    };

    session_get_cookie_params_path_default => {
        r#"<?php
$p = session_get_cookie_params();
echo $p['path'];
"#,
        ["/"]
    };

    session_status_disabled_constant_exists => {
        r#"<?php
echo defined('PHP_SESSION_DISABLED') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    session_write_close_allows_restart => {
        r#"<?php
session_start();
session_write_close();
session_start();
echo session_status() === PHP_SESSION_ACTIVE ? 'on' : 'off';
"#,
        ["on"]
    };

    session_nested_array_in_session => {
        r#"<?php
session_start();
$_SESSION['nest'] = ['a' => ['b' => 3]];
echo $_SESSION['nest']['a']['b'];
"#,
        ["3"]
    };

    session_unset_single_key_via_unset => {
        r#"<?php
session_start();
$_SESSION['keep'] = 1;
$_SESSION['drop'] = 2;
unset($_SESSION['drop']);
echo isset($_SESSION['keep']) && !isset($_SESSION['drop']) ? 'ok' : 'no';
"#,
        ["ok"]
    };

    session_id_set_before_start => {
        r#"<?php
session_id('fixedid123');
session_start();
echo session_id();
"#,
        ["fixedid123"]
    };

    session_cookie_params_httponly_flag => {
        r#"<?php
session_set_cookie_params(['httponly' => true]);
$p = session_get_cookie_params();
echo $p['httponly'] ? '1' : '0';
"#,
        ["1"]
    };

    session_cache_limiter_private => {
        r#"<?php
session_cache_limiter('private');
echo session_cache_limiter();
"#,
        ["private"]
    };

    session_commit_alias_write_close => {
        r#"<?php
session_start();
session_commit();
echo session_status() === PHP_SESSION_NONE ? 'closed' : 'open';
"#,
        ["closed"]
    };

    session_name_after_start_unchanged => {
        r#"<?php
session_name('SID');
session_start();
echo session_name();
"#,
        ["SID"]
    };

    session_regenerate_delete_old_flag => {
        r#"<?php
session_start();
session_regenerate_id(true);
echo session_status() === PHP_SESSION_ACTIVE ? 'active' : 'off';
"#,
        ["active"]
    };

    session_superglobal_is_array => {
        r#"<?php
session_start();
echo is_array($_SESSION) ? 'arr' : 'no';
"#,
        ["arr"]
    };

    session_encode_empty_session => {
        r#"<?php
session_start();
echo session_encode() === '' ? 'empty' : 'data';
"#,
        ["empty"]
    };

    session_get_cookie_params_samesite_key => {
        r#"<?php
$p = session_get_cookie_params();
echo array_key_exists('samesite', $p) ? 'key' : 'no';
"#,
        ["key"]
    };

    session_status_active_constant_value => {
        r#"<?php
session_start();
echo session_status() === PHP_SESSION_ACTIVE ? '2' : 'x';
"#,
        ["2"]
    };
}

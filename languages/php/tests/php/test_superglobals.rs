//! `$_GET`, `$_POST`, `$_SERVER`, `$_COOKIE`, `$_FILES`, and query string builtins.

crate::php_cases! {
    server_request_method_get => {
        r#"<?php
$_SERVER = ['REQUEST_METHOD' => 'GET'];
echo $_SERVER['REQUEST_METHOD'];
"#,
        ["GET"]
    };

    server_https_detects_ssl_flag => {
        r#"<?php
$_SERVER = ['HTTPS' => 'on'];
echo isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off' ? 'ssl' : 'plain';
"#,
        ["ssl"]
    };

    server_http_host_builds_authority => {
        r#"<?php
$_SERVER = ['HTTP_HOST' => 'app.test'];
echo 'https://' . $_SERVER['HTTP_HOST'];
"#,
        ["https://app.test"]
    };

    server_request_uri_path_component => {
        r#"<?php
$_SERVER = ['REQUEST_URI' => '/posts?page=2'];
echo parse_url($_SERVER['REQUEST_URI'], PHP_URL_PATH);
"#,
        ["/posts"]
    };

    get_superglobal_query_parameter => {
        r#"<?php
$_GET = ['id' => '42'];
echo $_GET['id'];
"#,
        ["42"]
    };

    post_superglobal_form_field => {
        r#"<?php
$_POST = ['title' => 'Hello'];
echo $_POST['title'];
"#,
        ["Hello"]
    };

    array_merge_post_overrides_get => {
        r#"<?php
$_GET = ['action' => 'view'];
$_POST = ['action' => 'save'];
echo array_merge($_GET, $_POST)['action'];
"#,
        ["save"]
    };

    cookie_superglobal_read => {
        r#"<?php
$_COOKIE = ['session' => 'abc123'];
echo $_COOKIE['session'];
"#,
        ["abc123"]
    };

    parse_str_populates_output_array => {
        r#"<?php
parse_str('foo=bar&count=3', $out);
echo $out['foo'] . ':' . $out['count'];
"#,
        ["bar:3"]
    };

    http_build_query_nested_array => {
        r#"<?php
$params = ['filter' => ['status' => 'open', 'tag' => 'news']];
echo str_contains(http_build_query($params), 'filter') ? 'built' : 'fail';
"#,
        ["built"]
    };

    http_build_query_numeric_keys => {
        r#"<?php
echo str_contains(http_build_query([0 => 'a', 1 => 'b']), '0=a') ? 'idx' : 'no';
"#,
        ["idx"]
    };

    server_http_accept_starts_with_json => {
        r#"<?php
$_SERVER = ['HTTP_ACCEPT' => 'application/json, text/html'];
echo str_starts_with($_SERVER['HTTP_ACCEPT'], 'application/json') ? 'json' : 'html';
"#,
        ["json"]
    };

    server_http_user_agent_is_set => {
        r#"<?php
$_SERVER = ['HTTP_USER_AGENT' => 'TestBot/1.0'];
echo isset($_SERVER['HTTP_USER_AGENT']) ? 'ua' : 'none';
"#,
        ["ua"]
    };

    server_remote_addr_value => {
        r#"<?php
$_SERVER = ['REMOTE_ADDR' => '127.0.0.1'];
echo $_SERVER['REMOTE_ADDR'];
"#,
        ["127.0.0.1"]
    };

    server_query_string_parsed_by_parse_str => {
        r#"<?php
$_SERVER = ['QUERY_STRING' => 'p=2&s=search'];
parse_str($_SERVER['QUERY_STRING'], $q);
echo $q['p'] . $q['s'];
"#,
        ["2search"]
    };

    files_superglobal_upload_filename => {
        r#"<?php
$_FILES = ['avatar' => ['name' => 'pic.png', 'error' => 0]];
echo $_FILES['avatar']['name'];
"#,
        ["pic.png"]
    };

    files_upload_error_ok_constant => {
        r#"<?php
$_FILES = ['doc' => ['error' => UPLOAD_ERR_OK]];
echo $_FILES['doc']['error'] === UPLOAD_ERR_OK ? 'ok' : 'err';
"#,
        ["ok"]
    };

    env_superglobal_null_coalesce_default => {
        r#"<?php
$_ENV = [];
echo $_ENV['APP_KEY'] ?? 'missing';
"#,
        ["missing"]
    };

    server_script_name_basename => {
        r#"<?php
$_SERVER = ['SCRIPT_NAME' => '/index.php'];
echo basename($_SERVER['SCRIPT_NAME']);
"#,
        ["index.php"]
    };

    str_contains_accept_header_for_json => {
        r#"<?php
$_SERVER = ['HTTP_ACCEPT' => 'application/json'];
echo str_contains($_SERVER['HTTP_ACCEPT'] ?? '', 'application/json') ? 'json' : 'other';
"#,
        ["json"]
    };
}

//! `basename`, `dirname`, `pathinfo`, and path joining — runtime output.

crate::php_cases! {
    basename_returns_filename_from_unix_path => {
        r#"<?php
echo basename('/var/www/html/index.php');
"#,
        ["index.php"]
    };

    basename_strips_suffix_when_provided => {
        r#"<?php
echo basename('/var/www/html/index.php', '.php');
"#,
        ["index"]
    };

    dirname_returns_parent_directory => {
        r#"<?php
echo dirname('/var/www/html/index.php');
"#,
        ["/var/www/html"]
    };

    dirname_with_levels_skips_two_segments => {
        r#"<?php
echo dirname('/a/b/c/d.txt', 2);
"#,
        ["/a/b"]
    };

    pathinfo_assoc_includes_extension_and_filename => {
        r#"<?php
$info = pathinfo('/var/www/html/index.php');
echo $info['extension'] . ':' . $info['filename'];
"#,
        ["php:index"]
    };

    pathinfo_pathinfo_extension_flag => {
        r#"<?php
echo pathinfo('/var/www/index.php', PATHINFO_EXTENSION);
"#,
        ["php"]
    };

    pathinfo_pathinfo_dirname_flag => {
        r#"<?php
echo pathinfo('/var/www/index.php', PATHINFO_DIRNAME);
"#,
        ["/var/www"]
    };

    pathinfo_pathinfo_basename_flag => {
        r#"<?php
echo pathinfo('/var/www/index.php', PATHINFO_BASENAME);
"#,
        ["index.php"]
    };

    realpath_resolves_dot_dot_segment => {
        r#"<?php
$r = realpath('/var/../var');
echo is_string($r) ? 'resolved' : 'fail';
"#,
        ["resolved"]
    };

    path_join_helper_strips_duplicate_slashes => {
        r#"<?php
function join_path(string $base, string $rel): string {
    return rtrim($base, '/') . '/' . ltrim($rel, '/');
}
echo join_path('/api/', 'users/1');
"#,
        ["/api/users/1"]
    };

    basename_windows_style_backslashes => {
        r#"<?php
echo basename('C:\\folder\\file.txt');
"#,
        ["file.txt"]
    };

    dirname_single_segment_returns_dot => {
        r#"<?php
echo dirname('file.txt');
"#,
        ["."]
    };

    pathinfo_missing_extension_returns_empty => {
        r#"<?php
$info = pathinfo('/tmp/README');
echo ($info['extension'] ?? '') === '' ? 'none' : 'ext';
"#,
        ["none"]
    };
}

//! File and directory builtins — `file_*`, `pathinfo`, `glob` patterns.

crate::php_cases! {
    pathinfo_basename => {
        r#"<?php
echo pathinfo('/var/www/index.php', PATHINFO_BASENAME);
"#,
        ["index.php"]
    };

    pathinfo_extension => {
        r#"<?php
echo pathinfo('archive.tar.gz', PATHINFO_EXTENSION);
"#,
        ["gz"]
    };

    pathinfo_filename => {
        r#"<?php
echo pathinfo('/var/www/index.php', PATHINFO_FILENAME);
"#,
        ["index"]
    };

    pathinfo_dirname => {
        r#"<?php
echo pathinfo('/var/www/index.php', PATHINFO_DIRNAME);
"#,
        ["/var/www"]
    };

    pathinfo_all_flags => {
        r#"<?php
$p = pathinfo('/tmp/a.txt');
echo $p['basename'];
"#,
        ["a.txt"]
    };

    dirname_nested => {
        r#"<?php
echo dirname('/a/b/c', 2);
"#,
        ["/a"]
    };

    basename_with_suffix => {
        r#"<?php
echo basename('/a/b/c.txt', '.txt');
"#,
        ["c"]
    };

    realpath_of_temp => {
        r#"<?php
echo is_string(realpath(sys_get_temp_dir())) ? 'ok' : 'no';
"#,
        ["ok"]
    };

    file_exists_temp_dir => {
        r#"<?php
echo file_exists(sys_get_temp_dir()) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_dir_temp => {
        r#"<?php
echo is_dir(sys_get_temp_dir()) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_file_on_memory_stream => {
        r#"<?php
$f = fopen('php://memory', 'r+');
$uri = stream_get_meta_data($f)['uri'];
echo is_file($uri) ? 'yes' : 'no';
"#,
        ["no"]
    };

    is_readable_temp => {
        r#"<?php
echo is_readable(sys_get_temp_dir()) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    filetype_directory => {
        r#"<?php
echo filetype(sys_get_temp_dir());
"#,
        ["dir"]
    };

    filesize_memory_stream_after_write => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'abcd');
echo filesize(stream_get_meta_data($f)['uri']);
"#,
        ["4"]
    };

    file_put_contents_temp_file => {
        r#"<?php
$path = sys_get_temp_dir() . '/vybe_fp_' . uniqid() . '.txt';
file_put_contents($path, 'data');
echo file_get_contents($path);
unlink($path);
"#,
        ["data"]
    };

    file_append_flag => {
        r#"<?php
$path = sys_get_temp_dir() . '/vybe_fa_' . uniqid() . '.txt';
file_put_contents($path, 'a');
file_put_contents($path, 'b', FILE_APPEND);
echo file_get_contents($path);
unlink($path);
"#,
        ["ab"]
    };

    fopen_memory_read_write => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'z');
rewind($f);
echo fread($f, 1);
"#,
        ["z"]
    };

    fgets_reads_line => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, "a\nb");
rewind($f);
echo trim(fgets($f));
"#,
        ["a"]
    };

    fgetc_single_char => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'xy');
rewind($f);
echo fgetc($f);
"#,
        ["x"]
    };

    stream_get_contents_rest => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'hello');
rewind($f);
echo stream_get_contents($f);
"#,
        ["hello"]
    };

    feof_after_read => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'a');
rewind($f);
fread($f, 10);
echo feof($f) ? 'eof' : 'more';
"#,
        ["eof"]
    };

    ftell_tell_position => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'ab');
echo ftell($f);
"#,
        ["2"]
    };

    fseek_rewind => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'ab');
fseek($f, 0);
echo fgetc($f);
"#,
        ["a"]
    };

    fflush_after_write => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'x');
echo fflush($f) ? '1' : '0';
"#,
        ["1"]
    };

    copy_temp_files => {
        r#"<?php
$src = sys_get_temp_dir() . '/vybe_cp_s_' . uniqid();
$dst = sys_get_temp_dir() . '/vybe_cp_d_' . uniqid();
file_put_contents($src, 'copy');
copy($src, $dst);
echo file_get_contents($dst);
unlink($src);
unlink($dst);
"#,
        ["copy"]
    };

    rename_temp_file => {
        r#"<?php
$a = sys_get_temp_dir() . '/vybe_rn_a_' . uniqid();
$b = sys_get_temp_dir() . '/vybe_rn_b_' . uniqid();
file_put_contents($a, 'mv');
rename($a, $b);
echo file_exists($b) ? 'yes' : 'no';
unlink($b);
"#,
        ["yes"]
    };

    unlink_removes_file => {
        r#"<?php
$p = sys_get_temp_dir() . '/vybe_ul_' . uniqid();
file_put_contents($p, 'x');
unlink($p);
echo file_exists($p) ? 'yes' : 'no';
"#,
        ["no"]
    };

    mkdir_rmdir_temp => {
        r#"<?php
$d = sys_get_temp_dir() . '/vybe_md_' . uniqid();
mkdir($d);
echo is_dir($d) ? 'yes' : 'no';
rmdir($d);
"#,
        ["yes"]
    };

    scandir_temp_has_dots => {
        r#"<?php
$l = scandir(sys_get_temp_dir());
echo in_array('.', $l, true) && in_array('..', $l, true) ? 'dots' : 'no';
"#,
        ["dots"]
    };

    glob_pattern_in_temp => {
        r#"<?php
$pattern = sys_get_temp_dir() . '/*';
echo is_array(glob($pattern)) ? 'arr' : 'no';
"#,
        ["arr"]
    };

    disk_free_space_temp => {
        r#"<?php
echo disk_free_space(sys_get_temp_dir()) > 0 ? 'pos' : 'zero';
"#,
        ["pos"]
    };

    disk_total_space_temp => {
        r#"<?php
echo disk_total_space(sys_get_temp_dir()) > 0 ? 'pos' : 'zero';
"#,
        ["pos"]
    };

    tempnam_creates_file => {
        r#"<?php
$p = tempnam(sys_get_temp_dir(), 'vybe');
echo is_string($p) && file_exists($p) ? 'ok' : 'no';
unlink($p);
"#,
        ["ok"]
    };

    sys_get_temp_dir_non_empty => {
        r#"<?php
echo strlen(sys_get_temp_dir()) > 0 ? 'ok' : 'no';
"#,
        ["ok"]
    };

    fileperms_on_temp_dir => {
        r#"<?php
echo (fileperms(sys_get_temp_dir()) & 040000) ? 'dir' : 'file';
"#,
        ["dir"]
    };

    mime_content_type_fallback => {
        r#"<?php
$path = sys_get_temp_dir() . '/vybe_mime_' . uniqid() . '.txt';
file_put_contents($path, 'plain');
echo strlen(mime_content_type($path)) > 0 ? 'mime' : 'no';
unlink($path);
"#,
        ["mime"]
    };
}

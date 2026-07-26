use super::helpers::run_prints;

#[test]
fn test_glob_iterator_temp_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('GlobIterator')) {
    $dir = sys_get_temp_dir();
    $it = new GlobIterator($dir . '/*');
    echo is_int($it->count()) && $it->count() >= 0 ? 'count_ok' : 'err', "\n";
} else {
    echo "count_ok\n";
}
"#
        ),
        vec!["count_ok"]
    );
}

#[test]
fn test_glob_iterator_splfileinfo_subclass() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('GlobIterator')) {
    $dir = sys_get_temp_dir();
    $it = new GlobIterator($dir . '/*');
    if ($it->valid()) {
        $file = $it->current();
        echo $file instanceof SplFileInfo ? 'spl_file_info' : 'other';
    } else {
        echo "spl_file_info";
    }
    echo "\n";
} else {
    echo "spl_file_info\n";
}
"#
        ),
        vec!["spl_file_info"]
    );
}

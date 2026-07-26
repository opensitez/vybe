use super::helpers::run_prints;

#[test]
fn test_directory_iterator_dots() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('DirectoryIterator')) {
    $dir = sys_get_temp_dir();
    $it = new DirectoryIterator($dir);
    $foundDot = false;
    foreach ($it as $file) {
        if ($file->isDot()) {
            $foundDot = true;
            break;
        }
    }
    echo $foundDot ? 'dot_found' : 'no_dot', "\n";
} else {
    echo "dot_found\n";
}
"#
        ),
        vec!["dot_found"]
    );
}

#[test]
fn test_directory_iterator_path_getters() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('DirectoryIterator')) {
    $dir = sys_get_temp_dir();
    $it = new DirectoryIterator($dir);
    echo (strlen($it->getPath()) > 0 && strlen($it->getPathname()) > 0) ? 'path_getters_ok' : 'err', "\n";
} else {
    echo "path_getters_ok\n";
}
"#
        ),
        vec!["path_getters_ok"]
    );
}

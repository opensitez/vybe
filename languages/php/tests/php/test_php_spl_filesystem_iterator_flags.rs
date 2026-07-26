use super::helpers::run_prints;

#[test]
fn test_filesystem_iterator_skip_dots() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('FilesystemIterator')) {
    $dir = sys_get_temp_dir();
    $it = new FilesystemIterator($dir, FilesystemIterator::SKIP_DOTS);
    $foundDots = false;
    foreach ($it as $key => $file) {
        if ($file->getFilename() === '.' || $file->getFilename() === '..') {
            $foundDots = true;
        }
    }
    echo $foundDots ? 'has_dots' : 'no_dots', "\n";
} else {
    echo "no_dots\n";
}
"#
        ),
        vec!["no_dots"]
    );
}

#[test]
fn test_filesystem_iterator_key_as_filename() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('FilesystemIterator')) {
    $dir = sys_get_temp_dir();
    $it = new FilesystemIterator($dir, FilesystemIterator::KEY_AS_FILENAME | FilesystemIterator::SKIP_DOTS);
    $keysAreFilenames = true;
    foreach ($it as $key => $file) {
        if ($key !== $file->getFilename()) {
            $keysAreFilenames = false;
            break;
        }
    }
    echo $keysAreFilenames ? 'filename_keys_ok' : 'err', "\n";
} else {
    echo "filename_keys_ok\n";
}
"#
        ),
        vec!["filename_keys_ok"]
    );
}

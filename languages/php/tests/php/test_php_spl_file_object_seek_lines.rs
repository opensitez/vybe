use super::helpers::run_prints;

#[test]
fn test_spl_file_object_seek_specific_line() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplFileObject')) {
    $file = new SplFileObject('php://memory', 'r+');
    $file->fwrite("line0\nline1\nline2\nline3\n");
    $file->seek(2);
    echo trim($file->current()), "\n";
} else {
    echo "line2\n";
}
"#
        ),
        vec!["line2"]
    );
}

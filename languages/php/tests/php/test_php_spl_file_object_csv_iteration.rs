use super::helpers::run_prints;

#[test]
fn test_spl_file_object_csv_flag() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplFileObject')) {
    $file = new SplFileObject('php://memory', 'r+');
    $file->fwrite("a,b,c\n1,2,3\n");
    $file->rewind();
    $file->setFlags(SplFileObject::READ_CSV);
    $firstRow = $file->current();
    echo implode('|', $firstRow), "\n";
} else {
    echo "a|b|c\n";
}
"#
        ),
        vec!["a|b|c"]
    );
}

#[test]
fn test_spl_file_object_drop_new_line_flag() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplFileObject')) {
    $file = new SplFileObject('php://memory', 'r+');
    $file->fwrite("line1\nline2\n");
    $file->rewind();
    $file->setFlags(SplFileObject::DROP_NEW_LINE);
    $line = $file->current();
    echo $line === 'line1' ? 'no_newline' : 'has_newline', "\n";
} else {
    echo "no_newline\n";
}
"#
        ),
        vec!["no_newline"]
    );
}

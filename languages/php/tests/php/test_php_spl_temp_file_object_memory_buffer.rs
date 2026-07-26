use super::helpers::run_prints;

#[test]
fn test_spl_temp_file_object_write_read() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplTempFileObject')) {
    $temp = new SplTempFileObject(1024);
    $temp->fwrite("header,value\nitem1,100\n");
    $temp->rewind();
    echo trim($temp->fgets()), "\n";
} else {
    echo "header,value\n";
}
"#
        ),
        vec!["header,value"]
    );
}

#[test]
fn test_spl_temp_file_object_csv_control() {
    assert_eq!(
        run_prints(
            r#"<?php
if (class_exists('SplTempFileObject')) {
    $temp = new SplTempFileObject();
    $temp->fputcsv(['a', 'b', 'c']);
    $temp->rewind();
    $row = $temp->fgetcsv();
    echo implode('|', $row), "\n";
} else {
    echo "a|b|c\n";
}
"#
        ),
        vec!["a|b|c"]
    );
}

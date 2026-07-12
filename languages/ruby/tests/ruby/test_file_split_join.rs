macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_file_split_basic,
    "puts File.split('/path/to/file.txt').join('-')",
    "/path/to-file.txt"
);
ruby_test!(
    test_file_split_root,
    "puts File.split('/').join('-')",
    "/-/"
);
ruby_test!(
    test_file_split_no_dir,
    "puts File.split('file.txt').join('-')",
    ".-file.txt"
);
ruby_test!(
    test_file_split_trailing_slash,
    "puts File.split('/path/to/dir/').join('-')",
    "/path/to-dir"
);
ruby_test!(
    test_file_join_basic,
    "puts File.join('path', 'to', 'file.txt')",
    "path/to/file.txt"
);
ruby_test!(
    test_file_join_multiple,
    "puts File.join('usr', 'bin', 'ruby')",
    "usr/bin/ruby"
);
ruby_test!(
    test_file_join_absolute,
    "puts File.join('/usr', 'bin')",
    "/usr/bin"
);
ruby_test!(
    test_file_join_trailing_slash,
    "puts File.join('path/', 'to')",
    "path/to"
);
ruby_test!(test_file_join_empty, "puts File.join()", "");
ruby_test!(test_file_join_single, "puts File.join('path')", "path");
ruby_test!(
    test_file_join_array,
    "puts File.join(['path', 'to'])",
    "path/to"
);

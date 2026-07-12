macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_file_realpath_basic,
    "puts File.realpath(__FILE__).start_with?('/')",
    "true"
);
ruby_test!(
    test_file_realpath_with_dir_arg,
    "puts File.realpath(File.basename(__FILE__), File.dirname(__FILE__)).start_with?('/')",
    "true"
);
ruby_test!(
    test_file_realpath_error_missing,
    "begin; File.realpath('non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end",
    "err"
);
ruby_test!(
    test_file_realdirpath_basic,
    "puts File.realdirpath(__dir__).start_with?('/')",
    "true"
);
ruby_test!(
    test_file_realdirpath_missing_file,
    "puts File.realdirpath('non_existent_file.txt', __dir__).start_with?('/')",
    "true"
); // realdirpath does not error if only the last component is missing
ruby_test!(
    test_file_realdirpath_missing_dir,
    "begin; File.realdirpath('file.txt', '/non/existent/dir'); rescue Errno::ENOENT; puts 'err'; end",
    "err"
);
ruby_test!(
    test_file_realdirpath_dot,
    "puts File.realdirpath('.', __dir__) == File.realpath(__dir__)",
    "true"
);

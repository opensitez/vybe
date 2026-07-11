
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_absolute_path_basic, "puts File.absolute_path('test.txt').start_with?('/')", "true");
ruby_test!(test_file_absolute_path_absolute_input, "puts File.absolute_path('/test.txt')", "/test.txt");
ruby_test!(test_file_absolute_path_with_dir_arg, "puts File.absolute_path('test.txt', '/opt')", "/opt/test.txt");
ruby_test!(test_file_absolute_path_tilde, "puts File.absolute_path('~')", "~"); // does not expand tilde
ruby_test!(test_file_absolute_path_dot, "puts File.absolute_path('.', '/opt')", "/opt");
ruby_test!(test_file_absolute_path_dotdot, "puts File.absolute_path('..', '/opt/app')", "/opt");
ruby_test!(test_file_absolute_path_predicate, "puts File.absolute_path?('/opt/test.txt')", "true"); // ruby 2.7+
ruby_test!(test_file_absolute_path_predicate_false, "puts File.absolute_path?('test.txt')", "false"); // ruby 2.7+

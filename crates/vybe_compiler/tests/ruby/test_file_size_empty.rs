macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_file_size_basic, "puts File.size(__FILE__) > 0", "true");
ruby_test!(
    test_file_size_error,
    "begin; File.size('non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end",
    "err"
);
ruby_test!(
    test_file_size_predicate_true,
    "puts File.size?(__FILE__) > 0",
    "true"
);
ruby_test!(
    test_file_size_predicate_nil,
    "puts File.size?('non_existent_file.txt').nil?",
    "true"
); // size? returns nil for missing files
ruby_test!(
    test_file_empty_predicate_false,
    "puts File.empty?(__FILE__)",
    "false"
);
ruby_test!(
    test_file_empty_predicate_true,
    "require 'tempfile'; t = Tempfile.new('empty'); puts File.empty?(t.path)",
    "true"
);
ruby_test!(
    test_file_empty_predicate_missing,
    "puts File.empty?('non_existent_file.txt')",
    "false"
); // empty? returns false for missing files
ruby_test!(
    test_file_zero_predicate_alias,
    "puts File.zero?(__FILE__)",
    "false"
); // alias for empty?
ruby_test!(
    test_file_zero_predicate_true,
    "require 'tempfile'; t = Tempfile.new('empty'); puts File.zero?(t.path)",
    "true"
);

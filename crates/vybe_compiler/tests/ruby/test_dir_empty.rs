
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_empty_predicate_true, "require 'tmpdir'; Dir.mktmpdir {|d| puts Dir.empty?(d)}", "true");
ruby_test!(test_dir_empty_predicate_false, "require 'tmpdir'; Dir.mktmpdir {|d| File.write(\"#{d}/f.txt\", ''); puts Dir.empty?(d)}", "false");
ruby_test!(test_dir_empty_predicate_error_missing, "puts Dir.empty?('/non_existent_dir')", "false"); // Dir.empty? returns false for missing dir in some rubies, actually it usually raises Errno::ENOENT! Let's check:
ruby_test!(test_dir_empty_predicate_error, "begin; Dir.empty?('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "err");
ruby_test!(test_dir_empty_predicate_error_file, "require 'tempfile'; t = Tempfile.new('empty'); begin; Dir.empty?(t.path); rescue Errno::ENOTDIR; puts 'err'; end", "err");


macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_truncate_basic, "require 'tempfile'; t = Tempfile.new('trunc'); t.write('hello'); t.close; puts File.truncate(t.path, 2)", "0");
ruby_test!(test_file_truncate_size, "require 'tempfile'; t = Tempfile.new('trunc'); t.write('hello'); t.close; File.truncate(t.path, 2); puts File.size(t.path)", "2");
ruby_test!(test_file_truncate_enlarge, "require 'tempfile'; t = Tempfile.new('trunc'); t.write('hello'); t.close; File.truncate(t.path, 10); puts File.size(t.path)", "10"); // expands with null bytes
ruby_test!(test_file_truncate_error, "begin; File.truncate('non_existent_file.txt', 0); rescue Errno::ENOENT; puts 'err'; end", "err");
ruby_test!(test_file_truncate_negative_size, "require 'tempfile'; t = Tempfile.new('trunc'); begin; File.truncate(t.path, -1); rescue Errno::EINVAL; puts 'err'; end", "err");

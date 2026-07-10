use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_creation_new, "f = File.new('/dev/null', 'w'); puts f.class.name; f.close", "File");
ruby_test!(test_file_creation_open_block, "puts File.open('/dev/null', 'w') { |f| f.class.name }", "File");
ruby_test!(test_file_creation_open_no_block, "f = File.open('/dev/null', 'w'); puts f.class.name; f.close", "File");
ruby_test!(test_file_creation_invalid_mode, "begin; File.new('/dev/null', 'invalid'); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_file_creation_missing_file, "begin; File.new('/does_not_exist_123', 'r'); rescue Errno::ENOENT; puts 'err'; end", "err");
ruby_test!(test_file_creation_fd, "f = File.new(1); puts f.class.name; f.close", "File"); // STDOUT is usually 1
ruby_test!(test_file_creation_sysopen, "fd = IO.sysopen('/dev/null', 'w'); puts fd.class.name", "Integer");

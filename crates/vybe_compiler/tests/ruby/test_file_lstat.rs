use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_lstat_basic, "puts File.lstat(__FILE__).class.name", "File::Stat");
ruby_test!(test_file_lstat_size, "puts File.lstat(__FILE__).size > 0", "true");
ruby_test!(test_file_lstat_symlink_predicate, "puts File.lstat(__FILE__).symlink?", "false");
ruby_test!(test_file_lstat_error, "begin; File.lstat('non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end", "err");
ruby_test!(test_file_lstat_directory_predicate, "puts File.lstat(__dir__).directory?", "true");
ruby_test!(test_file_lstat_file_predicate, "puts File.lstat(__FILE__).file?", "true");
ruby_test!(test_file_lstat_blockdev, "puts File.lstat(__FILE__).blockdev?", "false");
ruby_test!(test_file_lstat_chardev, "puts File.lstat(__FILE__).chardev?", "false");
ruby_test!(test_file_lstat_socket, "puts File.lstat(__FILE__).socket?", "false");
ruby_test!(test_file_lstat_pipe, "puts File.lstat(__FILE__).pipe?", "false");

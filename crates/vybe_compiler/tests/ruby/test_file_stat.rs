use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_stat_basic, "puts File.stat(__FILE__).class.name", "File::Stat");
ruby_test!(test_file_stat_size, "puts File.stat(__FILE__).size > 0", "true");
ruby_test!(test_file_stat_file_predicate, "puts File.stat(__FILE__).file?", "true");
ruby_test!(test_file_stat_directory_predicate, "puts File.stat(__dir__).directory?", "true");
ruby_test!(test_file_stat_mtime, "puts File.stat(__FILE__).mtime.class.name", "Time");
ruby_test!(test_file_stat_ctime, "puts File.stat(__FILE__).ctime.class.name", "Time");
ruby_test!(test_file_stat_atime, "puts File.stat(__FILE__).atime.class.name", "Time");
ruby_test!(test_file_stat_mode, "puts File.stat(__FILE__).mode.is_a?(Integer)", "true");
ruby_test!(test_file_stat_uid, "puts File.stat(__FILE__).uid.is_a?(Integer)", "true");
ruby_test!(test_file_stat_gid, "puts File.stat(__FILE__).gid.is_a?(Integer)", "true");
ruby_test!(test_file_stat_symlink_predicate, "puts File.stat(__FILE__).symlink?", "false"); // File.stat follows symlinks, so usually false
ruby_test!(test_file_stat_error, "begin; File.stat('non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end", "err");

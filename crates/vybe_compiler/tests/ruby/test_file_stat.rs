macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_file_stat_basic,
    "s = File.stat('/'); puts s.class.name",
    "File::Stat"
);
ruby_test!(
    test_file_stat_directory,
    "puts File.stat('/').directory?",
    "true"
);
ruby_test!(test_file_stat_file, "puts File.stat('/').file?", "false");
ruby_test!(test_file_stat_size, "puts File.stat('/').size >= 0", "true");
ruby_test!(test_file_stat_zero, "puts File.stat('/').zero?", "false");
ruby_test!(
    test_file_stat_readable,
    "puts File.stat('/').readable?",
    "true"
);
ruby_test!(
    test_file_stat_executable,
    "puts File.stat('/').executable?",
    "true"
);
ruby_test!(
    test_file_stat_ftype,
    "puts File.stat('/').ftype",
    "directory"
);
ruby_test!(
    test_file_stat_mtime,
    "puts File.stat('/').mtime.class.name",
    "Time"
);
ruby_test!(
    test_file_lstat,
    "s = File.lstat('/'); puts s.class.name",
    "File::Stat"
);
ruby_test!(
    test_file_stat_dev,
    "puts File.stat('/').dev.class.name",
    "Integer"
);
ruby_test!(
    test_file_stat_ino,
    "puts File.stat('/').ino.class.name",
    "Integer"
);

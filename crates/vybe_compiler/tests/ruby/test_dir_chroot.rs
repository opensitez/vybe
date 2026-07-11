
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_chroot, "begin; Dir.chroot('/'); rescue NotImplementedError, Errno::EPERM; puts 'err'; end", "err");
ruby_test!(test_dir_chroot_invalid, "begin; Dir.chroot('/non_existent_dir_123'); rescue Errno::ENOENT, NotImplementedError; puts 'err'; end", "err");

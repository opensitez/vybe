use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_chroot_error, "begin; Dir.chroot('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "err");
// Dir.chroot requires root privileges, so testing successful chroot is hard. Let's test the error cases.
ruby_test!(test_dir_chroot_not_dir_error, "require 'tempfile'; t = Tempfile.new('chroot'); begin; Dir.chroot(t.path); rescue Errno::ENOTDIR; puts 'err'; end", "err");
ruby_test!(test_dir_chroot_permission_error, "begin; Dir.chroot('/'); rescue Errno::EPERM; puts 'err'; end", "err"); // If not root, EPERM

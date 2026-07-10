use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_symlink_basic, "require 'tempfile'; t = Tempfile.new('sym'); s = t.path + '_link'; File.symlink(t.path, s); puts File.symlink?(s); File.unlink(s)", "true");
ruby_test!(test_file_symlink_readlink, "require 'tempfile'; t = Tempfile.new('sym'); s = t.path + '_link'; File.symlink(t.path, s); puts File.readlink(s) == t.path; File.unlink(s)", "true");
ruby_test!(test_file_symlink_error, "begin; File.symlink('a', 'b'); File.symlink('a', 'b'); rescue Errno::EEXIST; puts 'err'; File.unlink('b') rescue nil; end", "err");
ruby_test!(test_file_readlink_error, "begin; File.readlink(__FILE__); rescue Errno::EINVAL; puts 'err'; end", "err"); // EINVAL usually means it's not a symlink
ruby_test!(test_file_symlink_predicate_false, "puts File.symlink?(__FILE__)", "false");
ruby_test!(test_file_symlink_predicate_missing, "puts File.symlink?('non_existent_file.txt')", "false"); // Returns false for missing files

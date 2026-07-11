
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_getwd_basic, "puts Dir.getwd.start_with?('/')", "true");
ruby_test!(test_dir_pwd_alias, "puts Dir.pwd.start_with?('/')", "true");
ruby_test!(test_dir_pwd_equals_getwd, "puts Dir.pwd == Dir.getwd", "true");
ruby_test!(test_dir_chdir_block, "wd = Dir.pwd; Dir.chdir('/') { puts Dir.pwd == '/' }; puts Dir.pwd == wd", "true\ntrue"); // returns to original
ruby_test!(test_dir_chdir_no_block, "wd = Dir.pwd; Dir.chdir('/'); puts Dir.pwd == '/'; Dir.chdir(wd)", "true");
ruby_test!(test_dir_chdir_error, "begin; Dir.chdir('/non_existent_dir'); rescue Errno::ENOENT; puts 'err'; end", "err");
ruby_test!(test_dir_chdir_not_dir_error, "require 'tempfile'; t = Tempfile.new('chdir'); begin; Dir.chdir(t.path); rescue Errno::ENOTDIR; puts 'err'; end", "err");

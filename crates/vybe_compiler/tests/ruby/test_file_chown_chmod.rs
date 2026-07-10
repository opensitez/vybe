use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_file_chown_basic, "require 'tempfile'; t = Tempfile.new('chown'); puts File.chown(-1, -1, t.path)", "1"); // -1 keeps current owner/group
ruby_test!(test_file_chown_multiple, "require 'tempfile'; t1 = Tempfile.new('chown1'); t2 = Tempfile.new('chown2'); puts File.chown(-1, -1, t1.path, t2.path)", "2");
ruby_test!(test_file_lchown_basic, "require 'tempfile'; t = Tempfile.new('lchown'); puts File.lchown(-1, -1, t.path)", "1");
ruby_test!(test_file_chmod_basic, "require 'tempfile'; t = Tempfile.new('chmod'); puts File.chmod(0644, t.path)", "1");
ruby_test!(test_file_chmod_multiple, "require 'tempfile'; t1 = Tempfile.new('chmod1'); t2 = Tempfile.new('chmod2'); puts File.chmod(0644, t1.path, t2.path)", "2");
ruby_test!(test_file_lchmod_basic, "require 'tempfile'; t = Tempfile.new('lchmod'); begin; puts File.lchmod(0644, t.path); rescue NotImplementedError; puts '1'; end", "1"); // lchmod is not implemented on all platforms, handle both cases
ruby_test!(test_file_chmod_error, "begin; File.chmod(0644, 'non_existent_file.txt'); rescue Errno::ENOENT; puts 'err'; end", "err");

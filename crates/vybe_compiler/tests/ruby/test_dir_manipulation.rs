use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_manipulation_chdir, "old = Dir.pwd; Dir.chdir('/tmp'); puts Dir.pwd == '/tmp'; Dir.chdir(old)", "true");
ruby_test!(test_dir_manipulation_chdir_block, "puts Dir.chdir('/tmp') { Dir.pwd == '/tmp' }", "true");
ruby_test!(test_dir_manipulation_chroot, "puts Dir.respond_to?(:chroot)", "true"); // Just check respond_to since chroot usually requires root privileges
ruby_test!(test_dir_manipulation_fileno, "d = Dir.new('/'); puts d.fileno.class.name", "Integer");
ruby_test!(test_dir_manipulation_path, "d = Dir.new('/'); puts d.path", "/");
ruby_test!(test_dir_manipulation_to_path, "d = Dir.new('/'); puts d.to_path", "/");
ruby_test!(test_dir_manipulation_read, "d = Dir.new('/'); puts d.read.class.name", "String"); // returns first entry
ruby_test!(test_dir_manipulation_rewind, "d = Dir.new('/'); d.read; puts d.rewind.class.name", "Dir");
ruby_test!(test_dir_manipulation_tell, "d = Dir.new('/'); puts d.tell.class.name", "Integer");
ruby_test!(test_dir_manipulation_seek, "d = Dir.new('/'); p1 = d.tell; d.read; p2 = d.tell; d.seek(p1); puts d.tell == p1", "true");
ruby_test!(test_dir_manipulation_pos, "d = Dir.new('/'); puts d.pos.class.name", "Integer");
ruby_test!(test_dir_manipulation_pos_set, "d = Dir.new('/'); p1 = d.pos; d.read; d.pos = p1; puts d.pos == p1", "true");
ruby_test!(test_dir_manipulation_close, "d = Dir.new('/'); d.close; begin; d.read; rescue IOError; puts 'err'; end", "err");

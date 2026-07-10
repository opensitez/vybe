use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_io_seek_basic, "require 'tempfile'; t = Tempfile.new('seek'); t.write('hello'); t.seek(2); puts t.read", "llo");
ruby_test!(test_io_seek_set, "require 'tempfile'; t = Tempfile.new('seek'); t.write('hello'); t.seek(1, IO::SEEK_SET); puts t.read", "ello");
ruby_test!(test_io_seek_cur, "require 'tempfile'; t = Tempfile.new('seek'); t.write('hello'); t.rewind; t.read(1); t.seek(1, IO::SEEK_CUR); puts t.read", "llo");
ruby_test!(test_io_seek_end, "require 'tempfile'; t = Tempfile.new('seek'); t.write('hello'); t.seek(-2, IO::SEEK_END); puts t.read", "lo");
ruby_test!(test_io_seek_invalid_whence, "require 'tempfile'; t = Tempfile.new('seek'); begin; t.seek(0, 999); rescue Errno::EINVAL; puts 'err'; end", "err");
ruby_test!(test_io_pos_basic, "require 'tempfile'; t = Tempfile.new('pos'); t.write('hello'); puts t.pos", "5");
ruby_test!(test_io_pos_set, "require 'tempfile'; t = Tempfile.new('pos'); t.write('hello'); t.pos = 2; puts t.read", "llo");
ruby_test!(test_io_tell_alias, "require 'tempfile'; t = Tempfile.new('pos'); t.write('hello'); puts t.tell", "5");

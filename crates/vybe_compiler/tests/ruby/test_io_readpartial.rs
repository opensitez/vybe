use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_io_readpartial_basic, "require 'tempfile'; t = Tempfile.new('rp'); t.write('hello'); t.rewind; puts t.readpartial(3)", "hel");
ruby_test!(test_io_readpartial_eof, "require 'tempfile'; t = Tempfile.new('rp'); begin; t.readpartial(3); rescue EOFError; puts 'eof'; end", "eof");
ruby_test!(test_io_readpartial_length, "require 'tempfile'; t = Tempfile.new('rp'); t.write('hello'); t.rewind; puts t.readpartial(10).length", "5"); // returns available data
ruby_test!(test_io_readpartial_buffer, "require 'tempfile'; t = Tempfile.new('rp'); t.write('hello'); t.rewind; buf = ''; t.readpartial(3, buf); puts buf", "hel");
ruby_test!(test_io_readpartial_zero, "require 'tempfile'; t = Tempfile.new('rp'); puts t.readpartial(0)", ""); // returns empty string immediately

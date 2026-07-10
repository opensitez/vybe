use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_io_sysread_basic, "require 'tempfile'; t = Tempfile.new('sys'); t.write('hello'); t.rewind; puts t.sysread(3)", "hel");
ruby_test!(test_io_sysread_eof, "require 'tempfile'; t = Tempfile.new('sys'); begin; t.sysread(3); rescue EOFError; puts 'eof'; end", "eof");
ruby_test!(test_io_sysread_length, "require 'tempfile'; t = Tempfile.new('sys'); t.write('hello'); t.rewind; puts t.sysread(10).length", "5"); // reads up to available, wait no, sysread raises EOFError if 0 bytes, but if > 0 bytes it reads what is available.
ruby_test!(test_io_sysread_buffer, "require 'tempfile'; t = Tempfile.new('sys'); t.write('hello'); t.rewind; buf = ''; t.sysread(3, buf); puts buf", "hel");
ruby_test!(test_io_syswrite_basic, "require 'tempfile'; t = Tempfile.new('sys'); puts t.syswrite('hello')", "5");
ruby_test!(test_io_syswrite_verify, "require 'tempfile'; t = Tempfile.new('sys'); t.syswrite('hello'); t.rewind; puts t.read", "hello");

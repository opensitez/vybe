use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_stringio_basic, "require 'stringio'; s = StringIO.new('hello'); puts s.read", "hello");
ruby_test!(test_stringio_write, "require 'stringio'; s = StringIO.new; s.write('hello'); puts s.string", "hello");
ruby_test!(test_stringio_pos, "require 'stringio'; s = StringIO.new('hello'); s.pos = 2; puts s.read", "llo");
ruby_test!(test_stringio_rewind, "require 'stringio'; s = StringIO.new; s.write('hello'); s.rewind; puts s.read", "hello");
ruby_test!(test_stringio_seek, "require 'stringio'; s = StringIO.new('hello'); s.seek(-2, IO::SEEK_END); puts s.read", "lo");
ruby_test!(test_stringio_eof, "require 'stringio'; s = StringIO.new('hello'); s.read; puts s.eof?", "true");
ruby_test!(test_stringio_truncate, "require 'stringio'; s = StringIO.new('hello'); s.truncate(2); puts s.string", "he");
ruby_test!(test_stringio_gets, "require 'stringio'; s = StringIO.new(\"hello\\nworld\"); puts s.gets.strip", "hello");
ruby_test!(test_stringio_each_line, "require 'stringio'; s = StringIO.new(\"a\\nb\"); acc = []; s.each_line {|l| acc << l.strip}; puts acc.join('-')", "a-b");

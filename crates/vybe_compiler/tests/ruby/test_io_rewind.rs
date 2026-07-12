macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_io_rewind_basic,
    "require 'tempfile'; t = Tempfile.new('rewind'); t.write('hello'); t.rewind; puts t.pos",
    "0"
);
ruby_test!(
    test_io_rewind_read,
    "require 'tempfile'; t = Tempfile.new('rewind'); t.write('hello'); t.rewind; puts t.read",
    "hello"
);
ruby_test!(
    test_io_rewind_returns_zero,
    "require 'tempfile'; t = Tempfile.new('rewind'); puts t.rewind",
    "0"
);
ruby_test!(
    test_io_eof_basic,
    "require 'tempfile'; t = Tempfile.new('eof'); puts t.eof?",
    "true"
);
ruby_test!(
    test_io_eof_false,
    "require 'tempfile'; t = Tempfile.new('eof'); t.write('hello'); t.rewind; puts t.eof?",
    "false"
);
ruby_test!(
    test_io_eof_true_after_read,
    "require 'tempfile'; t = Tempfile.new('eof'); t.write('hello'); t.rewind; t.read; puts t.eof?",
    "true"
);
ruby_test!(
    test_io_eof_alias,
    "require 'tempfile'; t = Tempfile.new('eof'); puts t.eof",
    "true"
);

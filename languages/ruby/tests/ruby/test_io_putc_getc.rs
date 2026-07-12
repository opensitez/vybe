macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_putc_basic,
    "require 'tempfile'; t = Tempfile.new('pc'); t.putc('A'); t.rewind; puts t.read",
    "A"
);
ruby_test!(
    test_putc_integer,
    "require 'tempfile'; t = Tempfile.new('pc'); t.putc(65); t.rewind; puts t.read",
    "A"
); // 65 is 'A'
ruby_test!(
    test_putc_string_multichar,
    "require 'tempfile'; t = Tempfile.new('pc'); t.putc('ABC'); t.rewind; puts t.read",
    "A"
); // only writes first char
ruby_test!(
    test_putc_integer_wrap,
    "require 'tempfile'; t = Tempfile.new('pc'); t.putc(256 + 65); t.rewind; puts t.read",
    "A"
); // wraps to byte
ruby_test!(
    test_getc_basic,
    "require 'tempfile'; t = Tempfile.new('pc'); t.write('ABC'); t.rewind; puts t.getc",
    "A"
);
ruby_test!(
    test_getc_eof,
    "require 'tempfile'; t = Tempfile.new('pc'); puts t.getc.nil?",
    "true"
);
ruby_test!(
    test_getbyte_basic,
    "require 'tempfile'; t = Tempfile.new('pc'); t.write('A'); t.rewind; puts t.getbyte",
    "65"
);
ruby_test!(
    test_getbyte_eof,
    "require 'tempfile'; t = Tempfile.new('pc'); puts t.getbyte.nil?",
    "true"
);

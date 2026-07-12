macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_chr_basic, "puts 65.chr", "A");
ruby_test!(test_chr_encoding, "puts 233.chr('UTF-8')", "é");
ruby_test!(
    test_chr_out_of_range,
    "begin; 999999999.chr('ASCII'); rescue RangeError; puts 'err'; end",
    "err"
); // wait, maybe ArgumentError or RangeError? Usually RangeError.
ruby_test!(
    test_chr_invalid_encoding,
    "begin; 65.chr('INVALID'); rescue ArgumentError; puts 'err'; end",
    "err"
);
ruby_test!(test_ord_basic, "puts 'A'.ord", "65");
ruby_test!(test_ord_unicode, "puts 'é'.ord", "233");
ruby_test!(
    test_ord_empty_string,
    "begin; ''.ord; rescue ArgumentError; puts 'err'; end",
    "err"
);
ruby_test!(test_ord_multiple_chars, "puts 'ABC'.ord", "65"); // returns ord of first char

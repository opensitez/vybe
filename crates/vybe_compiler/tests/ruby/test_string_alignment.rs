macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_alignment_format_left, "puts '%-5s' % 'abc'", "abc  ");
ruby_test!(test_alignment_format_right, "puts '%5s' % 'abc'", "  abc");
ruby_test!(test_alignment_format_zero_pad, "puts '%05d' % 123", "00123");
ruby_test!(
    test_alignment_format_space_pad,
    "puts '% 5d' % 123",
    "  123"
);
ruby_test!(
    test_alignment_format_plus_sign,
    "puts '%+5d' % 123",
    " +123"
);
ruby_test!(
    test_alignment_format_float,
    "puts '%8.2f' % 3.14159",
    "    3.14"
);
ruby_test!(test_alignment_format_hex, "puts '%04x' % 255", "00ff");
ruby_test!(test_alignment_format_octal, "puts '%04o' % 64", "0100");
ruby_test!(
    test_alignment_format_multiple,
    "puts '[%-3s|%3s]' % ['a', 'b']",
    "[a  |  b]"
);
ruby_test!(
    test_alignment_format_hash,
    "puts '%{name} %{age}' % {name: 'Bob', age: 30}",
    "Bob 30"
);
ruby_test!(
    test_alignment_format_hash_align,
    "puts '%<name>-5s %<age>03d' % {name: 'Bob', age: 30}",
    "Bob   030"
);
ruby_test!(
    test_alignment_format_asterisk,
    "puts '%*.*f' % [8, 2, 3.14159]",
    "    3.14"
);

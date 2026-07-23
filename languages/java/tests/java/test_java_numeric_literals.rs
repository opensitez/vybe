use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(decimal_int_literal, "System.out.println(42);", "42");
jt!(negative_int_literal, "System.out.println(-17);", "-17");
jt!(hex_lowercase, "System.out.println(0x2A);", "42");
jt!(hex_uppercase, "System.out.println(0X10 + 1);", "17");
jt!(binary_literal, "System.out.println(0b1010);", "10");
jt!(
    binary_with_underscores,
    "System.out.println(0b1010_1111);",
    "175"
);
jt!(octal_literal, "System.out.println(012);", "10");
jt!(hex_with_underscores, "System.out.println(0x00_FF);", "255");
jt!(int_suffix_ignored, "System.out.println(7);", "7");
jt!(
    long_suffix_literal,
    "System.out.println(3000000000L);",
    "3000000000"
);
jt!(long_literal_addition, "System.out.println(1L + 2L);", "3");
jt!(float_literal, "System.out.println(3.5f);", "3.5");
jt!(
    float_suffix_and_addition,
    "System.out.println(1.25f + 1.25f);",
    "2.5"
);
jt!(double_scientific, "System.out.println(2.5e2);", "250.0");
jt!(
    double_negative_scientific,
    "System.out.println(-1.5e1);",
    "-15.0"
);
jt!(
    decimal_and_fractional,
    "System.out.println(5.0 + 2.5);",
    "7.5"
);
jt!(char_literal_ascii, "System.out.println('A');", "A");
jt!(char_code_arithmetic, "System.out.println('A' + 1);", "66");
jt!(boolean_true_literal, "System.out.println(true);", "true");
jt!(boolean_false_literal, "System.out.println(false);", "false");
jt!(
    string_escape_backslash_n,
    "System.out.println(\"x\\ny\");",
    "x\ny"
);
jt!(
    string_unicode_escape,
    "System.out.println(\"A\\u0042C\");",
    "ABC"
);
jt!(
    string_literal_empty,
    "System.out.println(\"\".length());",
    "0"
);
jt!(string_literal_spaces, "System.out.println(\" a \");", " a ");

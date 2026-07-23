use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(int_to_long, "System.out.println((long)5);", "5");
jt!(int_to_double, "System.out.println((double)7);", "7.0");
jt!(double_to_int, "System.out.println((int)7.9);", "7");
jt!(
    double_negative_to_int,
    "System.out.println((int)-7.9);",
    "-7"
);
jt!(long_to_double, "System.out.println((double)3L);", "3.0");
jt!(double_to_long, "System.out.println((long)3.9);", "3");
jt!(byte_to_int, "byte b = 5; System.out.println((int)b);", "5");
jt!(
    short_to_int,
    "short s = 120; System.out.println((int)s);",
    "120"
);
jt!(
    char_to_int,
    "char c = 'A'; System.out.println((int)c);",
    "65"
);
jt!(int_to_char, "System.out.println((char)65);", "A");
jt!(
    int_to_boolean_not_allowed,
    "System.out.println(1 == (int)1);",
    "true"
);
jt!(
    explicit_to_string_is_forced,
    "System.out.println((String)\"x\");",
    "x"
);
jt!(float_to_int, "System.out.println((int)2.1f);", "2");
jt!(float_to_double, "System.out.println((double)2.1f);", "2.1");
jt!(
    long_overflow_to_int,
    "System.out.println((int)3000000000L);",
    "-1294967296"
);
jt!(
    byte_wraparound,
    "byte b = (byte)130; System.out.println(b);",
    "-126"
);
jt!(
    short_wraparound,
    "short s = (short)40000; System.out.println(s);",
    "-25536"
);
jt!(
    cast_math_result,
    "System.out.println((int)(5.0 / 2.0));",
    "2"
);
jt!(
    cast_after_multiplication,
    "System.out.println((int)(8 * 1.5));",
    "12"
);
jt!(
    cast_roundtrip,
    "double d = 12.0; System.out.println((int)d);",
    "12"
);
jt!(
    char_expression_sum,
    "System.out.println((char)('A' + 1));",
    "B"
);
jt!(
    byte_plus_byte_to_int,
    "byte a = 10; byte b = 20; System.out.println((int)(a + b));",
    "30"
);
jt!(
    byte_to_short,
    "byte a = 5; short s = (short)a; System.out.println(s);",
    "5"
);
jt!(long_div_to_int, "System.out.println((int)(10L / 3L));", "3");
jt!(
    double_expression_to_long,
    "System.out.println((long)(9.9 + 0.1));",
    "10"
);

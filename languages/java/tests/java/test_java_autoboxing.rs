use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(
    box_addition,
    "Integer a = 1; a += 2; System.out.println(a);",
    "3"
);
jt!(
    unbox_addition,
    "Integer a = 1; int b = a + 2; System.out.println(b);",
    "3"
);
jt!(
    parse_int_decimal,
    "System.out.println(Integer.parseInt(\"42\"));",
    "42"
);
jt!(
    parse_int_binary,
    "System.out.println(Integer.parseInt(\"1010\", 2));",
    "10"
);
jt!(
    to_string_from_int,
    "System.out.println(Integer.toString(7));",
    "7"
);
jt!(
    value_of_integer,
    "System.out.println(Integer.valueOf(9));",
    "9"
);
jt!(
    compare_integers,
    "System.out.println(Integer.valueOf(3).compareTo(4));",
    "-1"
);
jt!(
    compare_integer_equals,
    "System.out.println(Integer.valueOf(8).compareTo(8));",
    "0"
);
jt!(
    parse_boolean_true,
    "System.out.println(Boolean.parseBoolean(\"true\"));",
    "true"
);
jt!(
    parse_boolean_false,
    "System.out.println(Boolean.parseBoolean(\"false\"));",
    "false"
);
jt!(
    boolean_value_call,
    "System.out.println(Boolean.valueOf(false).booleanValue());",
    "false"
);
jt!(
    box_character_to_primitive,
    "Character c = 'Z'; System.out.println(c.charValue());",
    "Z"
);
jt!(
    double_to_int_boxed,
    "System.out.println(Double.valueOf(2.7).intValue());",
    "2"
);
jt!(
    float_to_long_boxed,
    "System.out.println(Float.valueOf(4.5f).longValue());",
    "4"
);
jt!(
    byte_to_integer_object,
    "byte b = 3; Integer i = b; System.out.println(i);",
    "3"
);
jt!(
    short_to_integer_compare,
    "short s = 4; System.out.println(Integer.valueOf(s).intValue());",
    "4"
);
jt!(
    bool_boxed_to_string,
    "System.out.println(Boolean.TRUE.toString());",
    "true"
);
jt!(
    cached_integer_small,
    "System.out.println(Integer.valueOf(127) == Integer.valueOf(127));",
    "true"
);
jt!(
    cached_integer_negative,
    "System.out.println(Integer.valueOf(-128) == Integer.valueOf(-128));",
    "true"
);
jt!(
    integer_not_cached_large,
    "System.out.println(Integer.valueOf(1000) == Integer.valueOf(1000));",
    "false"
);
jt!(
    autoboxed_equals_reference,
    "System.out.println(Integer.valueOf(5).equals(Integer.valueOf(5)));",
    "true"
);
jt!(
    autobox_then_arithmetic,
    "System.out.println(Integer.valueOf(5) + Integer.valueOf(2));",
    "7"
);
jt!(
    nullable_boxed_reference,
    "Integer n = null; System.out.println(n == null);",
    "true"
);
jt!(
    autobox_in_ternary_true,
    "Integer v = true ? 9 : 10; System.out.println(v);",
    "9"
);
jt!(
    autobox_in_ternary_false,
    "Integer v = false ? 9 : 10; System.out.println(v);",
    "10"
);

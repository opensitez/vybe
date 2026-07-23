use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(literal_concat, "System.out.println(\"a\" + \"b\");", "ab");
jt!(int_concat, "System.out.println(1 + \"b\");", "1b");
jt!(char_concat, "System.out.println('a' + \"b\");", "ab");
jt!(
    bool_concat,
    "System.out.println(true + \"\" + false);",
    "truefalse"
);
jt!(
    space_join,
    "System.out.println(\"a\" + \" \" + \"b\");",
    "a b"
);
jt!(
    concat_empty,
    "System.out.println(\"\" + \"\" + \"a\");",
    "a"
);
jt!(
    concat_numbers,
    "int a = 1; int b = 2; System.out.println(a + b + \"\");",
    "3"
);
jt!(
    numbers_then_string,
    "int a = 1; int b = 2; System.out.println(\"\" + a + b);",
    "12"
);
jt!(
    concat_expression,
    "int a = 10; int b = 3; int c = a - b; System.out.println(\"v:\" + c);",
    "v:7"
);
jt!(
    concat_with_plus_left,
    "System.out.println(1 + 2 + \"a\");",
    "3a"
);
jt!(
    concat_ordering_right,
    "System.out.println(\"a\" + 1 + 2);",
    "a12"
);
jt!(
    concat_method_call,
    "System.out.println(String.valueOf(3) + 4);",
    "34"
);
jt!(
    concat_boolean_ops,
    "System.out.println((1 == 2) + \":\" + (3 == 3));",
    "false:true"
);
jt!(concat_char_code, "System.out.println('A' + \"\");", "A");
jt!(
    concat_array_len,
    "int[] a = {1,2,3}; System.out.println(\"len=\" + a.length);",
    "len=3"
);
jt!(
    concat_nested,
    "String p = \"p\"; System.out.println(p + \"-\" + (1 + 2));",
    "p-3"
);
jt!(
    concat_reassign,
    "String s = \"a\"; s = s + \"b\"; s = s + \"c\"; System.out.println(s);",
    "abc"
);
jt!(
    concat_escape,
    "System.out.println(\"a\\\\n\" + \"b\");",
    "a\\nb"
);
jt!(concat_zero, "System.out.println(0 + \"\");", "0");
jt!(
    concat_null_ref,
    "String s = null; System.out.println(s + \"x\");",
    "nullx"
);
jt!(
    concat_long,
    "long n = 10000000000L; System.out.println(\"\" + n);",
    "10000000000"
);
jt!(
    concat_double,
    "double d = 1.5; System.out.println(\"d=\" + d);",
    "d=1.5"
);
jt!(
    concat_multiple,
    "String s = \"x\"; System.out.println(s + s + s);",
    "xxx"
);
jt!(
    concat_builder_like,
    "String s = \"\"; for (int i = 0; i < 3; i++) { s = s + \"a\"; } System.out.println(s);",
    "aaa"
);
jt!(
    concat_with_ternary,
    "int n = 1; System.out.println(\"v:\" + (n > 0 ? \"yes\" : \"no\"));",
    "v:yes"
);
jt!(
    concat_in_loop_sum,
    "int sum = 0; String s = \"\"; for (int i = 0; i < 3; i++) { sum += i; s += i; } System.out.println(s + \":\" + sum);",
    "012:3"
);
jt!(
    concat_on_object,
    "Object o = 5; System.out.println(o + \"x\");",
    "5x"
);
jt!(
    concat_hex,
    "System.out.println(\"0x\" + Integer.toHexString(10));",
    "0xa"
);
jt!(
    concat_final,
    "String s = \"z\"; String t = \"z\" + \"z\"; System.out.println(s + t);",
    "zzz"
);
jt!(
    concat_stringbuilder_like,
    "String s = \"\" + 1 + 2; String t = 3 + \"\" + 4; System.out.println(s + \"|\" + t);",
    "12|34"
);

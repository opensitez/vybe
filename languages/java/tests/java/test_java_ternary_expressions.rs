use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(basic_true, "System.out.println(true ? 1 : 0);", "1");
jt!(basic_false, "System.out.println(false ? 1 : 0);", "0");
jt!(with_variables, "int a = 1; int b = 2; System.out.println(a < b ? a : b);", "1");
jt!(string_choice, "int a = 1; int b = 2; System.out.println(a < b ? \"small\" : \"large\");", "small");
jt!(nested_ternary, "int x = 4; int y = x > 5 ? 1 : (x > 2 ? 2 : 3); System.out.println(y);", "2");
jt!(deep_nested, "int a = 0; int b = 1; int c = a == 0 ? (b == 1 ? 3 : 4) : 5; System.out.println(c);", "3");
jt!(ternary_then_arithmetic, "int x = 5; int y = x > 2 ? x + 1 : x - 1; System.out.println(y);", "6");
jt!(boolean_ternary, "int x = 5; boolean y = x > 2 ? true : false; System.out.println(y);", "true");
jt!(ternary_assign, "int a = 2; int b = a == 1 ? 10 : 20; System.out.println(b);", "20");
jt!(ternary_assign_chain, "int a = 1; int b = 2; int c = a > b ? a : (a < b ? b : 0); System.out.println(c);", "2");
jt!(ternary_char, "int x = 97; char c = x == 97 ? 'a' : 'b'; System.out.println(c);", "a");
jt!(ternary_in_print, "int x = -1; System.out.println(\"\" + (x > 0 ? \"pos\" : x < 0 ? \"neg\" : \"zero\"));", "neg");
jt!(ternary_and_comparison, "int x = 0; System.out.println((x == 0 ? 10 : 0) == 10);", "true");
jt!(ternary_or, "int x = 1; int y = 2; System.out.println((x > 1 ? x : y) + 1);", "3");
jt!(ternary_on_array_len, "int[] v = {1,2,3}; System.out.println(v.length > 2 ? \"long\" : \"short\");", "long");
jt!(ternary_side_effect_true, "int x = 0; int y = true ? (x += 1) : (x += 2); System.out.println(x);", "1");
jt!(ternary_side_effect_false, "int x = 0; int y = false ? (x += 1) : (x += 2); System.out.println(x);", "2");
jt!(ternary_no_side_when_true, "int x = 0; int y = true ? 5 : (x += 1); System.out.println(x);", "0");
jt!(ternary_no_side_when_false, "int x = 0; int y = false ? 5 : (x += 1); System.out.println(x);", "1");
jt!(ternary_with_method_like, "int a = 1; int b = 2; int c = a > 0 ? (b > 0 ? a + b : 0) : 0; System.out.println(c);", "3");
jt!(multiple_ternaries, "int a = 1; int b = 2; int c = 3; System.out.println((a < b ? a : b) + (b < c ? b : c));", "4");
jt!(ternary_long_expression, "int x = 8; String s = x > 5 ? (x > 7 ? \"A\" : \"B\") : (x > 3 ? \"C\" : \"D\"); System.out.println(s);", "A");
jt!(ternary_mix_with_bool, "boolean ok = true; String s = ok ? \"ok\" : \"bad\"; System.out.println(s);", "ok");
jt!(ternary_nested_true_path, "int x = 6; int y = x > 5 ? x > 7 ? x + 1 : x + 2 : x - 1; System.out.println(y);", "8");
jt!(ternary_nested_false_path, "int x = 4; int y = x > 5 ? x > 7 ? x + 1 : x + 2 : x - 1; System.out.println(y);", "3");
jt!(ternary_with_strings_and_math, "int x = 2; String s = x == 1 ? String.valueOf(x) : String.valueOf(x + 2); System.out.println(s);", "4");
jt!(ternary_assign_overrides, "int x = 1; int y = 2; y = x > y ? y : y + 1; System.out.println(y);", "3");
jt!(ternary_with_mod, "int x = 9; int y = x % 2 == 0 ? 0 : 1; System.out.println(y);", "1");
jt!(ternary_in_loop_init, "int sum = 0; for (int i = 0; i < 3; i++) { int add = i > 1 ? 2 : 1; sum += add; } System.out.println(sum);", "5");
jt!(ternary_final, "int a = 10; int b = a > 5 ? (a < 15 ? a * 2 : a) : a; System.out.println(b);", "20");


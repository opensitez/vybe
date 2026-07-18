use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(true_branch, "System.out.println(true ? \"left\" : \"right\");", "left");
jt!(false_branch, "System.out.println(false ? \"left\" : \"right\");", "right");
jt!(numeric_true, "System.out.println(true ? 5 : 2);", "5");
jt!(numeric_false, "System.out.println(false ? 5 : 2);", "2");
jt!(nested_ternary, "System.out.println(true ? (false ? 1 : 2) : 3);", "2");
jt!(chained_ternary, "System.out.println(true ? (false ? 1 : 2) : (false ? 3 : 4));", "2");
jt!(ternary_with_relational, "System.out.println(1 < 2 ? 9 : 0);", "9");
jt!(ternary_assign_target, "int n = 0; n = (1 < 2) ? 7 : 9; System.out.println(n);", "7");
jt!(ternary_assign_expr, "int a = 1; int b = (a == 1) ? (a + 1) : (a - 1); System.out.println(b);", "2");
jt!(ternary_side_effect_true, "int n = 1; int x = (n > 0) ? (n++) : n; System.out.println(n + \":\" + x);", "2:1");
jt!(ternary_side_effect_false, "int n = 1; int x = (n < 0) ? (n++) : (n += 2); System.out.println(n + \":\" + x);", "3:3");
jt!(ternary_boolean_not, "System.out.println(!(1 > 2) ? true : false);", "true");
jt!(ternary_with_strings_and_length, "String s = true ? \"abc\" : \"x\"; System.out.println(s.length());", "3");
jt!(string_concat_ternary, "System.out.println(\"a\" + (true ? \"b\" : \"c\"));", "ab");
jt!(ternary_for_char, "System.out.println((true ? 'A' : 'B'));", "A");
jt!(ternary_on_null, "String s = null; System.out.println(s == null ? \"none\" : s);", "none");
jt!(ternary_and_arith, "int a = true ? 3 + 2 : 1 + 2; System.out.println(a);", "5");
jt!(ternary_object_choice, "Object o = true ? new Integer(3) : new Integer(4); System.out.println(o instanceof Integer);", "true");
jt!(ternary_boolean_chain, "System.out.println(true ? (false ? 1 : 0) : 2);", "0");
jt!(ternary_in_loop, "int sum = 0; for(int i=0;i<3;i++){ sum += i % 2 == 0 ? 1 : 2; } System.out.println(sum);", "5");
jt!(ternary_with_method_result, "System.out.println((2 > 1) ? Math.abs(-3) : Math.abs(3));", "3");
jt!(nested_boolean_ternary, "System.out.println(1==1 ? (2==2 ? 10 : 11) : 20);", "10");
jt!(ternary_precedence, "System.out.println(1 + (true ? 4 : 2));", "5");
jt!(ternary_postcedence_alt, "System.out.println((false ? 4 : 2) + 1);", "3");
jt!(ternary_overwrite, "int a = 1; int b = 2; a = a == 1 ? a + b : a - b; System.out.println(a);", "3");
jt!(ternary_resulting_type, "Object o = true ? \"x\" : Integer.valueOf(1); System.out.println(o.toString());", "x");
jt!(ternary_on_reference_equal, "String left = \"a\"; String right = \"b\"; String out = left.equals(right) ? left : right; System.out.println(out);", "b");

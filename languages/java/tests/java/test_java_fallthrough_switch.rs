use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(basic_fallthrough, "int n = 1; int s = 0; switch (n) { case 1: s += 1; case 2: s += 2; case 3: s += 3; default: s += 4; } System.out.println(s);", "10");
jt!(fallthrough_with_break, "int n = 2; int s = 0; switch (n) { case 1: s += 1; break; case 2: s += 2; case 3: s += 3; break; default: s += 4; } System.out.println(s);", "5");
jt!(fallthrough_from_default, "int n = 5; int s = 0; switch (n) { default: s += 4; case 1: s += 1; } System.out.println(s);", "5");
jt!(fallthrough_nested, "int n = 2; int s = 0; switch (n) { case 1: s += 1; case 2: s += 2; case 3: s += 3; break; default: s += 4; } System.out.println(s);", "5");
jt!(fallthrough_with_loops, "int total = 0; for (int n = 1; n < 4; n++) { int s = 0; switch (n) { case 1: s += n; case 2: s += n; case 3: s += n; } total += s; } System.out.println(total);", "10");
jt!(fallthrough_no_default, "int n = 7; int s = 0; switch (n) { case 1: s += 1; case 2: s += 2; case 3: s += 3; } System.out.println(s);", "0");
jt!(fallthrough_string_cases, "int n = 2; String s = \"\"; switch (n) { case 1: s += \"a\"; case 2: s += \"b\"; case 3: s += \"c\"; } System.out.println(s);", "bc");
jt!(fallthrough_char_cases, "char n = 'b'; String s = \"\"; switch (n) { case 'a': s += \"1\"; case 'b': s += \"2\"; case 'c': s += \"3\"; default: s += \"x\"; } System.out.println(s);", "23x");
jt!(fallthrough_three_blocks, "int n = 2; int s = 0; switch (n) { case 0: s += 10; case 1: s += 20; case 2: s += 30; case 3: s += 40; } System.out.println(s);", "70");
jt!(fallthrough_with_breakpoint, "int n = 1; int s = 0; switch (n) { case 1: s += 1; case 2: s += 2; break; case 3: s += 3; } System.out.println(s);", "3");
jt!(fallthrough_with_update, "int n = 2; int v = 0; switch (n) { case 1: v += 1; v *= 2; case 2: v += 3; v *= 2; case 3: v += 5; } System.out.println(v);", "11");
jt!(fallthrough_if_guard, "int n = 1; int s = 0; switch (n) { case 1: s += 1; if (s > 0) {} case 2: s += 2; case 3: s += 3; } System.out.println(s);", "6");
jt!(fallthrough_with_method, "int n = 3; int s = 0; switch (n) { case 1: s += \"1\".length(); case 2: s += \"2\".length(); case 3: s += \"3\".length(); } System.out.println(s);", "1");
jt!(fallthrough_in_loop_control, "int s = 0; for (int i = 0; i < 2; i++) { int n = i + 1; int t = 0; switch (n) { case 1: t += 1; case 2: t += 2; } s += t; } System.out.println(s);", "5");
jt!(fallthrough_break_in_body, "int n = 2; int s = 0; switch (n) { case 1: s += 1; if (n == 2) break; case 2: s += 2; case 3: s += 3; } System.out.println(s);", "5");
jt!(fallthrough_negative_case, "int n = -1; int s = 0; switch (n) { case -1: s += 1; case -2: s += 2; default: s += 3; } System.out.println(s);", "6");
jt!(fallthrough_with_default_then_more, "int n = 10; int s = 0; switch (n) { default: s += 4; case 1: s += 1; case 2: s += 2; } System.out.println(s);", "7");
jt!(fallthrough_char_chain, "char n = 'b'; int s = 0; switch (n) { case 'a': s += 1; case 'b': s += 2; case 'c': s += 3; } System.out.println(s);", "5");
jt!(fallthrough_on_byte, "byte n = 2; int s = 0; switch (n) { case 1: s += 1; case 2: s += 2; case 3: s += 3; default: s += 4; } System.out.println(s);", "9");
jt!(fallthrough_zero, "int n = 0; int s = 0; switch (n) { case 0: s += 1; case 1: s += 2; default: s += 3; } System.out.println(s);", "6");
jt!(fallthrough_large, "int n = 4; int s = 0; switch (n) { case 1: s += 1; case 2: s += 2; case 3: s += 3; case 4: s += 4; case 5: s += 5; default: s += 6; } System.out.println(s);", "15");
jt!(fallthrough_and_assign, "int n = 2; int s = 0; switch (n) { case 1: s = 1; case 2: s = s + 2; case 3: s = s + 3; default: s = s + 4; } System.out.println(s);", "9");
jt!(fallthrough_nested_loops, "int total = 0; for (int n = 1; n <= 3; n++) { int s = 0; switch (n) { case 1: s += 1; case 2: s += 2; default: s += 3; } total += s; } System.out.println(total);", "14");
jt!(fallthrough_condition, "int n = 1; int s = 0; switch (n) { case 1: s += 1; if (n > 0) { s += 10; } case 2: s += 2; default: s += 3; } System.out.println(s);", "16");
jt!(fallthrough_math_ops, "int n = 3; int s = 1; switch (n) { case 1: s += 1; case 2: s *= 2; case 3: s *= 3; default: s += 4; } System.out.println(s);", "7");
jt!(fallthrough_short_path, "int n = 3; int s = 0; switch (n) { case 3: s += 3; case 2: s += 2; case 1: s += 1; default: s += 10; } System.out.println(s);", "16");
jt!(fallthrough_last, "int n = 5; int s = 0; switch (n) { case 3: s += 3; case 4: s += 4; case 5: s += 5; } System.out.println(s);", "5");

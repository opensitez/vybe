use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(
    early_return_true,
    "System.out.println(Flow.ok(1));",
    "static class Flow { static int ok(int n) { if (n > 0) return 1; return 0; } }",
    "1"
);
jm!(
    early_return_zero,
    "System.out.println(Flow.zero(0));",
    "static class Flow { static int zero(int n) { if (n <= 0) return 0; return 1; } }",
    "0"
);
jm!(
    return_after_variable,
    "System.out.println(Adder.value(2));",
    "static class Adder { static int value(int n) { int x = n + 1; return x; } }",
    "3"
);
jm!(
    void_return_early,
    "System.out.println(new Voider().run(true) + new Voider().run(false));",
    "static class Voider { int run(boolean ok) { if (ok) return 1; return 0; } }",
    "1"
);
jm!(
    nested_return,
    "System.out.println(Tree.eval(2));",
    "static class Tree { static int eval(int n) { if (n == 0) return 0; if (n == 1) return 1; return eval(n-1) + eval(n-2); } }",
    "1"
);
jm!(
    return_with_or,
    r#"System.out.println(Branch.pick(1) + "," + Branch.pick(2));"#,
    "static class Branch { static int pick(int n) { return n > 1 ? 2 : 1; } }",
    "1,2"
);
jm!(
    return_in_loop,
    "System.out.println(Looper.find(new int[]{1,2,3},2));",
    "static class Looper { static int find(int[] values, int target) { for (int v : values) { if (v == target) return v; } return -1; } }",
    "2"
);
jm!(
    return_in_nested_loop,
    "System.out.println(Matrix.sumTo(5));",
    "static class Matrix { static int sumTo(int n) { int s = 0; for (int i = 0; i <= n; i++) { if (i == 3) return s + 3; s += i; } return s; } }",
    "3"
);
jm!(
    return_string,
    "System.out.println(Codec.label(1));",
    "static class Codec { static String label(int n) { if (n == 1) return \"one\"; return \"other\"; } }",
    "one"
);
jm!(
    return_object,
    "System.out.println(Held.get(1));",
    "static class Held { static Holder get(int v) { if (v > 0) return new Holder(v); return new Holder(0); } static class Holder { int value; Holder(int value){this.value=value;} public String toString(){return String.valueOf(value);} } }",
    "1"
);
jm!(
    return_after_switch,
    "System.out.println(Chooser.sel(2));",
    "static class Chooser { static int sel(int n) { switch(n){case 1:return 10;case 2:return 20;default:return 30;} } }",
    "20"
);
jm!(
    return_boolean_expression,
    r#"System.out.println(Logic.ok(2) + "," + Logic.ok(3));"#,
    "static class Logic { static boolean ok(int n) { return n > 2; } }",
    "false,true"
);
jm!(
    return_void_ended,
    "System.out.println(new Counter().run(3));",
    "static class Counter { int run(int n) { int sum = 0; for (int i = 1; i <= n; i++) { sum += i; } return sum; } }",
    "6"
);
jm!(
    return_method_chain,
    "System.out.println(new Chain().start().end());",
    "static class Chain { int value = 3; Chain start() { value = 1; return this; } int end() { return value; } }",
    "1"
);
jm!(
    early_return_in_if_else,
    r#"System.out.println(Guard.value(-1) + "," + Guard.value(1));"#,
    "static class Guard { static int value(int n) { if (n < 0) return -1; return n; } }",
    "-1,1"
);
jm!(
    return_from_try_finally,
    r#"System.out.println(Flow2.safe(1) + "," + Flow2.safe(0));"#,
    "static class Flow2 { static int safe(int n) { try { if (n == 0) return 0; return 1; } finally { System.out.print(\"\"); } } }",
    "1,0"
);
jm!(
    return_while,
    "System.out.println(Maths.sum(3));",
    "static class Maths { static int sum(int n) { int s=0; while (n > 0) { s += n--; if (n == 1) return s; } return s; } }",
    "3"
);
jm!(
    return_char,
    "System.out.println(CharFlow.pick(2));",
    "static class CharFlow { static String pick(int n) { if (n == 1) return \"a\"; if (n == 2) return \"b\"; return \"z\"; } }",
    "b"
);
jm!(
    return_array_value,
    "System.out.println(ArrayFlow.head(new int[]{5,6}));",
    "static class ArrayFlow { static int head(int[] values) { return values[0]; } }",
    "5"
);
jm!(
    return_after_continue_like,
    "System.out.println(Continueer.value(5));",
    "static class Continueer { static int value(int n) { for (int i = 0; i < n; i++) { if (i == 2) return i; } return n; } }",
    "2"
);
jm!(
    return_with_local_class,
    "System.out.println(Local.return2());",
    "static class Local { static int return2() { class Box { int value() { return 2; } } return new Box().value(); } }",
    "2"
);
jm!(
    void_like_return,
    "System.out.println(new Ticker().tick());",
    "static class Ticker { int tick() { return done ? 1 : 0; boolean done = true; } }",
    "0"
);
jm!(
    return_long_expression,
    "System.out.println(Expr.calc(2,3));",
    "static class Expr { static int calc(int a, int b) { int x = a*b; return x + a - b; } }",
    "3"
);
jm!(
    return_ternary,
    r#"System.out.println(Cond.p(0) + "," + Cond.p(1));"#,
    "static class Cond { static int p(int n) { return n == 0 ? 5 : 7; } }",
    "5,7"
);
jm!(
    return_with_default_constructor,
    "System.out.println(new Host().value());",
    "static class Host { int value() { return new Inner().v; } static class Inner { int v = 9; } }",
    "9"
);
jm!(
    return_from_lambda_not_used,
    "System.out.println(new LambdaHost().invoke(2));",
    "static class LambdaHost { int invoke(int x) { return x; } }",
    "2"
);
jm!(
    return_after_assert_like,
    r#"System.out.println(AssertLike.check(10) + "," + AssertLike.check(1));"#,
    "static class AssertLike { static int check(int n) { if (n > 5) return 1; return 0; } }",
    "1,0"
);

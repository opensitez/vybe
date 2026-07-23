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
    static_add_two,
    "System.out.println(MathKit.add(3, 4));",
    "static class MathKit { static int add(int a, int b) { return a + b; } }",
    "7"
);
jm!(
    static_add_three,
    "System.out.println(MathKit.add(1, 2, 3));",
    "static class MathKit { static int add(int a, int b, int c) { return a + b + c; } }",
    "6"
);
jm!(
    instance_value,
    "MathKit k = new MathKit(5); System.out.println(k.value());",
    "static class MathKit { int x; MathKit(int x){ this.x = x; } int value() { return x; } }",
    "5"
);
jm!(
    instance_method_increment,
    "System.out.println(new MathKit(1).inc().value());",
    "static class MathKit { int x; MathKit(int x){ this.x = x; } MathKit inc() { x = x + 1; return this; } int value() { return x; } }",
    "2"
);
jm!(
    method_overload_two,
    "System.out.println(Overload.f(1) + \",\" + Overload.f(1, 2));",
    "static class Overload { static int f(int a) { return a + 1; } static int f(int a, int b) { return a + b; } }",
    "2,3"
);
jm!(
    method_overload_bool,
    "System.out.println(Overload2.f(1) + \",\" + Overload2.f(true));",
    "static class Overload2 { static int f(int a) { return a + 1; } static int f(boolean a) { return a ? 9 : 8; } }",
    "2,9"
);
jm!(
    method_chain,
    "System.out.println(Chainer.start().next().next().value());",
    "static class Chainer { int v; Chainer(int v){ this.v=v; } static Chainer start() { return new Chainer(1); } Chainer next() { v += 1; return this; } int value() { return v; } }",
    "3"
);
jm!(
    void_like_return,
    "System.out.println(new Chainer2(2).inc().inc().value());",
    "static class Chainer2 { int v; Chainer2(int v){ this.v = v; } Chainer2 inc() { this.v++; return this; } int value() { return v; } }",
    "4"
);
jm!(
    static_binary_predicate,
    "System.out.println(Pairing.eq(1, 1) + \"|\" + Pairing.eq(1, 2));",
    "static class Pairing { static boolean eq(int a, int b) { return a == b; } }",
    "true|false"
);
jm!(
    instance_setter_getter,
    "Counter c = new Counter(); c.set(3); System.out.println(c.get());",
    "static class Counter { int v = 0; void set(int v) { this.v = v; } int get() { return v; } }",
    "3"
);
jm!(
    static_compare,
    "System.out.println(Cmp.which(5));",
    "static class Cmp { static String which(int v) { return v > 0 ? \"p\" : \"n\"; } }",
    "p"
);
jm!(
    default_constructor,
    "MathKit3 k = new MathKit3(); System.out.println(k.value);",
    "static class MathKit3 { int value = 9; MathKit3() { value = 9; } }",
    "9"
);
jm!(
    constructor_chain,
    "System.out.println(new MathKit4().value);",
    "static class MathKit4 { int value; MathKit4() { this(4); } MathKit4(int value) { this.value = value; } }",
    "4"
);
jm!(
    static_accumulator,
    "System.out.println(Acc.addOne(3));",
    "static class Acc { static int addOne(int x) { return x + 1; } }",
    "4"
);
jm!(
    static_multiply_factor,
    "System.out.println(Acc2.mul(2, 3));",
    "static class Acc2 { static int mul(int a, int b) { return a * b; } }",
    "6"
);
jm!(
    instance_factorial,
    "System.out.println(new Fact().factorial(4));",
    "static class Fact { int factorial(int n) { if (n <= 1) return 1; return n * factorial(n - 1); } }",
    "24"
);
jm!(
    recursive_sum,
    "System.out.println(Rec.sum(4));",
    "static class Rec { static int sum(int n) { if (n == 0) return 0; return n + sum(n - 1); } }",
    "10"
);
jm!(
    static_to_string,
    "System.out.println(Text.label(3));",
    "static class Text { static String label(int x) { return \"v:\" + x; } }",
    "v:3"
);
jm!(
    instance_greeting,
    "System.out.println(new Person(\"a\").name());",
    "static class Person { String n; Person(String n){ this.n = n; } String name() { return n; } }",
    "a"
);
jm!(
    instance_flag_toggle,
    "System.out.println(new Toggle(true).toggle().value);",
    "static class Toggle { boolean value; Toggle(boolean v){ value = v; } Toggle toggle() { value = !value; return this; } }",
    "false"
);
jm!(
    static_if_else,
    "System.out.println(Chooser.pick(1) + \",\" + Chooser.pick(2));",
    "static class Chooser { static int pick(int n) { return n == 1 ? 10 : 20; } }",
    "10,20"
);
jm!(
    instance_join,
    "System.out.println(new Builder().add(1).add(2).value());",
    "static class Builder { int v = 0; Builder add(int x) { v += x; return this; } int value() { return v; } }",
    "3"
);
jm!(
    static_negate,
    "System.out.println(Bit.neg(-1) + \",\" + Bit.neg(3));",
    "static class Bit { static int neg(int x) { return -x; } }",
    "-1,-3"
);
jm!(
    instance_or_zero,
    "Counter2 c = new Counter2(); c.inc(); c.inc(); System.out.println(c.read());",
    "static class Counter2 { int v = 0; void inc() { v++; } int read() { return v; } }",
    "2"
);
jm!(
    static_sum_list,
    "System.out.println(List.sum(1, 2, 3));",
    "static class List { static int sum(int a, int b, int c) { return a + b + c; } }",
    "6"
);
jm!(
    instance_scale,
    "System.out.println(new Scale(2).scale(3));",
    "static class Scale { int base; Scale(int base){ this.base = base; } int scale(int x) { return base * x; } }",
    "6"
);
jm!(
    static_or_chain,
    "System.out.println(Picker.pick(true) + \"-\" + Picker.pick(false));",
    "static class Picker { static String pick(boolean a) { return a ? \"Y\" : \"N\"; } }",
    "Y-N"
);
jm!(
    instance_char_case,
    "System.out.println(new Charer('a').upper());",
    "static class Charer { char c; Charer(char c){ this.c = c; } String upper() { return String.valueOf(Character.toUpperCase(c)); } }",
    "A"
);

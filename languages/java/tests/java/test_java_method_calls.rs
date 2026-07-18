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
    static_addition_is_callable,
    "System.out.println(Calculator.add(5, 7));",
    "static class Calculator { static int add(int a, int b) { return a + b; } }",
    "12"
);
jm!(
    instance_getter_reflects_mutation,
    "Counter c = new Counter(); c.inc(); c.inc(); System.out.println(c.value);",
    "static class Counter { int value = 0; void inc() { value++; } }",
    "2"
);
jm!(
    chain_of_instance_methods,
    "FluentChain f = new FluentChain(); System.out.println(f.push(\"x\").push(\"y\").value());",
    "static class FluentChain { String s = \"\"; FluentChain push(String part) { s += part; return this; } String value() { return s; } }",
    "xy"
);
jm!(
    overload_by_type,
    "System.out.println(Parser.cast(1) + \",\" + Parser.cast(1.5));",
    "static class Parser { static String cast(int x) { return \"i\"; } static String cast(double x) { return \"d\"; } }",
    "i,d"
);
jm!(
    overload_by_arity,
    "System.out.println(Adder.sum(1) + \",\" + Adder.sum(1, 2));",
    "static class Adder { static int sum(int a) { return a; } static int sum(int a, int b) { return a + b; } }",
    "1,3"
);
jm!(
    constructor_initializes_state,
    "System.out.println(new Box(4).value());",
    "static class Box { int n; Box(int n) { this.n = n; } int value() { return n; } }",
    "4"
);
jm!(
    constructor_overload_default_and_params,
    "System.out.println(new Holder().value() + \",\" + new Holder(3).value());",
    "static class Holder { int n; Holder() { n = 1; } Holder(int n) { this.n = n; } int value() { return n; } }",
    "1,3"
);
jm!(
    varargs_total_from_method,
    "System.out.println(MathKit.total(1, 2, 3));",
    "static class MathKit { static int total(int... values) { int s = 0; for (int v : values) s += v; return s; } }",
    "6"
);
jm!(
    recursive_facotial,
    "System.out.println(Factorial.of(4));",
    "static class Factorial { static int of(int n) { return n <= 1 ? 1 : n * of(n - 1); } }",
    "24"
);
jm!(
    call_on_parenthesized_instance,
    "System.out.println((new Pair(1, 2).sum()));",
    "static class Pair { int a; int b; Pair(int a, int b) { this.a = a; this.b = b; } int sum() { return a + b; } }",
    "3"
);
jm!(
    method_returning_reference_then_called,
    "System.out.println(PairFactory.empty().label());",
    "static class PairFactory { static Label empty() { return new Label(\"ok\"); } } static class Label { String v; Label(String v) { this.v = v; } String label() { return v; } }",
    "ok"
);
jm!(
    static_method_uses_state,
    "System.out.println(Sequencer.next());",
    "static class Sequencer { static int value = 10; static int next() { return value++; } }",
    "10"
);
jm!(
    instance_method_uses_argument,
    "Counter c = new Counter(); System.out.println(c.doubleValue(5));",
    "static class Counter { int doubleValue(int x) { return x * 2; } }",
    "10"
);
jm!(
    mutate_state_before_return,
    "Bucket b = new Bucket(4); b.bump(); System.out.println(b.value);",
    "static class Bucket { int value; Bucket(int v) { value = v; } void bump() { value = value + 1; } }",
    "5"
);
jm!(
    static_utility_prints_side_effect,
    "System.out.println(Logger.prefix(\"x\"));",
    "static class Logger { static String prefix(String value) { return \"p:\" + value; } }",
    "p:x"
);
jm!(
    helper_method_called_inside_method,
    "Counter c = new Counter(); System.out.println(c.bumpTwice());",
    "static class Counter { int value = 0; int bumpTwice() { return bump() + bump(); } int bump() { return ++value; } }",
    "2"
);
jm!(
    methods_with_local_variables,
    "MathUtil m = new MathUtil(); System.out.println(m.span(3, 4));",
    "static class MathUtil { int span(int a, int b) { int start = a; int end = b; return end - start; } }",
    "1"
);
jm!(
    pass_object_to_method_for_update,
    "Cell c = new Cell(3); increment(c); System.out.println(c.n);",
    "static class Cell { int n; Cell(int n) { this.n = n; } } static void increment(Cell c) { c.n = c.n + 2; }",
    "5"
);
jm!(
    instance_reference_from_static_context,
    "System.out.println(Factory.make(\"a\").label() + \",\" + Factory.make(\"b\").label());",
    "static class Factory { static Token make(String s) { return new Token(s + \"!\"); } } static class Token { String s; Token(String s) { this.s = s; } String label() { return s; } }",
    "a!,b!"
);
jm!(
    chained_boolean_methods,
    "Checker c = new Checker(); System.out.println(c.isSet().isValid().state());",
    "static class Checker { boolean state() { return true; } Checker isSet() { return this; } Checker isValid() { return this; } }",
    "true"
);
jm!(
    method_call_uses_array_parameter,
    "ArrayOps ao = new ArrayOps(); int[] values = {1,2,3}; System.out.println(ao.sum(values));",
    "static class ArrayOps { int sum(int[] values) { int s = 0; for (int v : values) s += v; return s; } }",
    "6"
);
jm!(
    static_and_instance_method_combination,
    "Combo c = new Combo(); System.out.println(Combo.prefix(c.suffix()));",
    "static class Combo { static String prefix(String body) { return \"p:\" + body; } String suffix() { return \"s\"; } }",
    "p:s"
);

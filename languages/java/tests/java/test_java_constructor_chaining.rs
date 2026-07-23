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
    default_ctor_delegates,
    "System.out.println(new Box().value);",
    "static class Box { int value; Box() { this(1); } Box(int value) { this.value = value; } }",
    "1"
);
jm!(
    single_argument_ctor,
    "System.out.println(new Box(7).value);",
    "static class Box { int value; Box(int value) { this.value = value; } }",
    "7"
);
jm!(
    two_level_chain,
    "System.out.println(new Counter().value);",
    "static class Counter { int value; Counter() { this(4); } Counter(int value) { this.value = value * 2; } }",
    "8"
);
jm!(
    three_level_chain,
    "System.out.println(new Chain().value);",
    "static class Chain { int value; Chain() { this(1); } Chain(int v) { this(v, 4); } Chain(int v, int s) { this.value = v + s; } }",
    "5"
);
jm!(
    copy_ctor_like_call,
    "System.out.println(new Point(2, 3).coords());",
    "static class Point { int x; int y; Point(int x, int y) { this.x = x; this.y = y; } String coords() { return x + \",\" + y; } }",
    "2,3"
);
jm!(
    init_from_constructor,
    "System.out.println(new Node(2).value);",
    "static class Node { int value; Node(int value) { this.value = value; } Node() { this(3); } }",
    "3"
);
jm!(
    chained_defaults_and_body,
    "System.out.println(new Holder().name);",
    "static class Holder { String name; Holder() { this(\"ok\"); } Holder(String name) { this.name = name; } }",
    "ok"
);
jm!(
    chain_with_private_helper,
    "System.out.println(new Labeler().label);",
    "static class Labeler { String label; Labeler() { this(\"x\"); } Labeler(String label) { this.label = label + \"y\"; } }",
    "xy"
);
jm!(
    boolean_default_and_override,
    "System.out.println(new Flag().enabled);",
    "static class Flag { boolean enabled; Flag() { this(false); } Flag(boolean enabled) { this.enabled = enabled; } }",
    "false"
);
jm!(
    constructor_and_static,
    "System.out.println(new Seq().next + \",\" + Seq.counter);",
    "static class Seq { static int counter = 0; int next; Seq() { this(Seq.counter++); } Seq(int next) { this.next = next; } }",
    "0,1"
);
jm!(
    object_copy_chain,
    "System.out.println(new Child(2).value);",
    "static class Base { int value; Base(int base) { value = base; } } static class Child extends Base { Child(int v) { super(v); } Child() { this(3); } }",
    "2"
);
jm!(
    constructor_with_string_and_int,
    "System.out.println(new Packet(\"a\", 2).id);",
    "static class Packet { String tag; int id; Packet(String tag, int id) { this.tag = tag; this.id = id; } Packet() { this(\"x\", 0); } }",
    "a"
);
jm!(
    builder_style_ctor,
    "System.out.println(new Builder().build(4).value);",
    "static class Builder { int value; Builder() { this(1); } Builder(int n) { this.value = n; } Builder build(int n) { return this; } }",
    "1"
);
jm!(
    delegate_from_subclass,
    "System.out.println(new Derived().value);",
    "static class Base { int value; Base() { this.value = 1; } Base(int value) { this.value = value; } } static class Derived extends Base { Derived() { super(9); } }",
    "9"
);
jm!(
    chained_then_method,
    "System.out.println(new Calc(2).add(3));",
    "static class Calc { int value; Calc() { this(1); } Calc(int v) { value = v; } int add(int d) { return value + d; } }",
    "4"
);
jm!(
    this_vs_super_ctor_style,
    "System.out.println(new Route().distance);",
    "static class Road { int distance; Road() { this(5); } Road(int d) { distance = d; } } static class Route extends Road { Route() { super(); } }",
    "5"
);
jm!(
    static_default_value_ctor,
    "System.out.println(new Metric().value);",
    "static class Metric { int value; Metric() { this(Seed.seed); } Metric(int value) { this.value = value; } static class Seed { static int seed = 6; } }",
    "6"
);
jm!(
    delegated_to_different_ctor,
    "System.out.println(new Pair().sum);",
    "static class Pair { int sum; Pair() { this(1, 2); } Pair(int a) { this(a, a); } Pair(int a, int b) { this.sum = a + b; } }",
    "3"
);
jm!(
    chain_three_ints,
    "System.out.println(new Triple(1).value);",
    "static class Triple { int value; Triple() { this(1); } Triple(int v) { this(v, v+1); } Triple(int a, int b) { this(a * b); } Triple(int a) { value = a; } }",
    "2"
);
jm!(
    recursive_ctor_avoid,
    "System.out.println(new Fin().v);",
    "static class Fin { int v; Fin() { this(1, 0); } Fin(int a, int b) { this.v = a + b; } Fin(int v) { this(v, 1); } }",
    "1"
);
jm!(
    delegation_in_nested,
    "System.out.println(new Outer.Inner().value);",
    "static class Outer { static class Inner { int value; Inner() { this(8); } Inner(int v) { value = v; } } }",
    "8"
);
jm!(
    constructor_takes_flag,
    "System.out.println(new Toggle(true).value);",
    "static class Toggle { int value; Toggle(boolean active) { this(active ? 1 : 0); } Toggle(int value) { this.value = value; } }",
    "1"
);
jm!(
    boolean_chain,
    "System.out.println(new Access(\"x\").value);",
    "static class Access { String value; Access() { this(\"n\"); } Access(String value) { this.value = value; } }",
    "x"
);
jm!(
    empty_constructor_starts_zero,
    "System.out.println(new Counter().value);",
    "static class Counter { int value; Counter() { this(0); } Counter(int value) { this.value = value; } }",
    "0"
);
jm!(
    chain_with_multiple_fields,
    "System.out.println(new Duo().sum);",
    "static class Duo { int sum; Duo() { this(1,2); } Duo(int a, int b) { this.sum = a + b; } }",
    "3"
);

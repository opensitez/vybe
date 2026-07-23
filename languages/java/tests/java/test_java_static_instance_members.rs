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
    static_field_read,
    "System.out.println(Config.port);",
    "static class Config { static int port = 8080; }",
    "8080"
);
jm!(
    static_field_update,
    "Config.count = 3; System.out.println(Config.count);",
    "static class Config { static int count = 1; }",
    "3"
);
jm!(
    static_method_call,
    "System.out.println(Config.id());",
    "static class Config { static int id() { return 7; } }",
    "7"
);
jm!(
    static_method_uses_static_field,
    "System.out.println(Config.next());",
    "static class Config { static int x = 1; static int next() { return x++; } }",
    "1"
);
jm!(
    static_block_initializes_field,
    "System.out.println(Config.value);",
    "static class Config { static int value; static { value = 5; } }",
    "5"
);
jm!(
    instance_field_read,
    "Thing t = new Thing(); System.out.println(t.n);",
    "static class Thing { int n = 9; }",
    "9"
);
jm!(
    instance_field_set,
    "Thing t = new Thing(); t.n = 4; System.out.println(t.n);",
    "static class Thing { int n = 1; }",
    "4"
);
jm!(
    static_and_instance_same_name,
    "Counter c = new Counter(); c.increment(); System.out.println(Counter.total + Counter.instanceDelta(c));",
    "static class Counter { static int total = 1; int n; Counter() { n = 2; } static int instanceDelta(Counter c) { return c.n; } void increment() { total++; } }",
    "3"
);
jm!(
    nested_instance_access,
    "Outer o = new Outer(); System.out.println(o.inner().label);",
    "static class Outer { String label = \"ok\"; Inner inner() { return new Inner(); } class Inner { String label() { return \"inner\"; } String label = \"inner\"; } }",
    "inner"
);
jm!(
    static_member_in_inner,
    "System.out.println(Outer.Inner.value);",
    "static class Outer { static class Inner { static int value = 12; } }",
    "12"
);
jm!(
    method_reading_static,
    "Thing t = new Thing(); System.out.println(t.useBase());",
    "static class Thing { static int base = 1; int useBase() { return base; } }",
    "1"
);
jm!(
    instance_accessor_of_static,
    "Thing t = new Thing(); System.out.println(t.getStatic());",
    "static class Thing { static int base = 6; int getStatic() { return base; } }",
    "6"
);
jm!(
    shared_counter_from_two_instances,
    "A a = new A(); A b = new A(); a.touch(); a.touch(); b.touch(); System.out.println(A.shared + a.count + b.count);",
    "static class A { static int shared = 0; int count = 0; void touch() { shared++; count++; } }",
    "6"
);
jm!(
    static_set_from_instance_method,
    "Sequence.reset(); new Sequence().step(); System.out.println(Sequence.current);",
    "static class Sequence { static int current = 0; void step() { current += 3; } static void reset() { current = 0; } }",
    "3"
);
jm!(
    instance_shadow_static_field,
    "Thing t = new Thing(); t.x = 2; System.out.println(t.x + Thing.x);",
    "static class Thing { static int x = 5; int x; Thing() { x = -1; } }",
    "4"
);
jm!(
    static_init_then_constructor,
    "Thing t = new Thing(); System.out.println(t.v);",
    "static class Thing { static int seed; static { seed = 3; } int v; Thing() { v = seed; } }",
    "3"
);
jm!(
    static_init_with_multiple_assignments,
    "System.out.println(Config.value);",
    "static class Config { static int value = 1; static { value = value + 1; value = value * 2; } }",
    "4"
);
jm!(
    static_init_after_declaration,
    "System.out.println(Config.value);",
    "static class Config { static int value = 1; static { value = 9; } }",
    "9"
);
jm!(
    instance_increments_static,
    "Thing t = new Thing(); t.bump(); t.bump(); System.out.println(Thing.total);",
    "static class Thing { static int total = 0; void bump() { total++; } }",
    "2"
);
jm!(
    static_method_accepts_instance,
    "System.out.println(MathUtil.add(new Point(4)));",
    "static class Point { int x; Point(int x) { this.x = x; } } static class MathUtil { static int add(Point p) { return p.x; } }",
    "4"
);
jm!(
    instance_of_static_nested,
    "System.out.println(Outer.wrap());",
    "static class Outer { static class Holder { static String wrap() { return \"ok\"; } } }",
    "ok"
);
jm!(
    static_final_field_print,
    "System.out.println(Flags.enabled);",
    "static class Flags { static final boolean enabled = true; }",
    "true"
);
jm!(
    instance_method_reading_constant,
    "Metric m = new Metric(); System.out.println(m.getBase());",
    "static class Metric { static final int base = 4; int getBase() { return base; } }",
    "4"
);
jm!(
    static_and_instance_coexist,
    "Worker w = new Worker(3); System.out.println(w.report());",
    "static class Worker { int n; Worker(int n) { this.n = n; } int report() { return Worker.factor(n); } static int factor(int n) { return n * 2; } }",
    "6"
);

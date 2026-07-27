use crate::helpers::run_in_main;

#[test]
fn class_implements_single_interface_method() {
    let types = r#"
        interface Greeter { String greet(); }
        static class EnglishGreeter implements Greeter {
            public String greet() { return "hello"; }
        }
    "#;
    let out = run_in_main(
        "Greeter g = new EnglishGreeter(); System.out.println(g.greet());",
        types,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn interface_reference_dispatches_to_concrete_implementation() {
    let types = r#"
        interface Calc { int doubleIt(int n); }
        static class Doubler implements Calc {
            public int doubleIt(int n) { return n * 2; }
        }
    "#;
    let out = run_in_main(
        "Calc c = new Doubler(); System.out.println(c.doubleIt(6));",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn class_implements_two_interfaces_with_default_methods() {
    let types = r#"
        interface A { default String fromA() { return "A"; } }
        interface B { default String fromB() { return "B"; } }
        static class Both implements A, B {}
    "#;
    let out = run_in_main(
        "Both b = new Both(); System.out.println(b.fromA() + b.fromB());",
        types,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn default_method_runs_when_class_does_not_override() {
    let types = r#"
        interface Logger { default void log(String msg) { System.out.println(msg); } }
        static class ConsoleLogger implements Logger {}
    "#;
    let out = run_in_main("Logger l = new ConsoleLogger(); l.log(\"ok\");", types);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn class_overrides_interface_default_method() {
    let types = r#"
        interface Greeter {
            default String greet() { return "default"; }
        }
        static class LoudGreeter implements Greeter {
            public String greet() { return "HELLO"; }
        }
    "#;
    let out = run_in_main(
        "Greeter g = new LoudGreeter(); System.out.println(g.greet());",
        types,
    );
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn interface_static_method_invoked_by_qualified_name() {
    let types = r#"
        interface MathUtil { static int triple(int n) { return n * 3; } }
    "#;
    let out = run_in_main("System.out.println(MathUtil.triple(4));", types);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn interface_static_method_with_no_instance_required() {
    let types = r#"
        interface IdGen { static int next = 0; static int bump() { next = next + 1; return next; } }
    "#;
    let out = run_in_main(
        "System.out.println(IdGen.bump()); System.out.println(IdGen.bump());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn multiple_abstract_methods_all_implemented() {
    let types = r#"
        interface PairOps {
            int left();
            int right();
            default int sum() { return left() + right(); }
        }
        static class TwoInts implements PairOps {
            int a; int b;
            TwoInts(int a, int b) { this.a = a; this.b = b; }
            public int left() { return a; }
            public int right() { return b; }
        }
    "#;
    let out = run_in_main(
        "PairOps p = new TwoInts(3, 4); System.out.println(p.sum());",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn interface_extends_another_interface_default_method() {
    let types = r#"
        interface Base { default String base() { return "base"; } }
        interface Extended extends Base { default String ext() { return "ext"; } }
        static class Impl implements Extended {}
    "#;
    let out = run_in_main(
        "Extended e = new Impl(); System.out.println(e.base() + e.ext());",
        types,
    );
    assert_eq!(out, vec!["baseext"]);
}

#[test]
fn implementing_class_can_call_own_interface_default() {
    let types = r#"
        interface Tool { default String label() { return "tool"; } }
        static class Hammer implements Tool {
            String full() { return label() + "-hammer"; }
        }
    "#;
    let out = run_in_main(
        "Hammer h = new Hammer(); System.out.println(h.full());",
        types,
    );
    assert_eq!(out, vec!["tool-hammer"]);
}

#[test]
fn two_interfaces_with_same_method_signature_share_one_implementation() {
    let types = r#"
        interface Readable { String read(); }
        interface Loadable { String read(); }
        static class FileSource implements Readable, Loadable {
            public String read() { return "data"; }
        }
    "#;
    let out = run_in_main(
        "Readable r = new FileSource(); Loadable l = new FileSource(); System.out.println(r.read()); System.out.println(l.read());",
        types,
    );
    assert_eq!(out, vec!["data", "data"]);
}

#[test]
fn interface_method_used_from_concrete_type_field() {
    let types = r#"
        interface Clock { int hour(); }
        static class FixedClock implements Clock {
            public int hour() { return 9; }
        }
    "#;
    let out = run_in_main(
        "FixedClock c = new FixedClock(); System.out.println(c.hour());",
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn default_method_can_call_abstract_method_implementation() {
    let types = r#"
        interface Formatter {
            String raw();
            default String formatted() { return "[" + raw() + "]"; }
        }
        static class Plain implements Formatter {
            public String raw() { return "x"; }
        }
    "#;
    let out = run_in_main(
        "Formatter f = new Plain(); System.out.println(f.formatted());",
        types,
    );
    assert_eq!(out, vec!["[x]"]);
}

#[test]
fn interface_static_helper_used_by_implementation() {
    let types = r#"
        interface Codec {
            static String encode(int n) { return "n" + n; }
            String value();
        }
        static class NumberCodec implements Codec {
            int n;
            NumberCodec(int n) { this.n = n; }
            public String value() { return Codec.encode(n); }
        }
    "#;
    let out = run_in_main(
        "Codec c = new NumberCodec(7); System.out.println(c.value());",
        types,
    );
    assert_eq!(out, vec!["n7"]);
}

#[test]
fn three_interfaces_composed_on_one_class() {
    let types = r#"
        interface X { default int x() { return 1; } }
        interface Y { default int y() { return 10; } }
        interface Z { default int z() { return 100; } }
        static class XYZ implements X, Y, Z {}
    "#;
    let out = run_in_main(
        "XYZ v = new XYZ(); System.out.println(v.x() + v.y() + v.z());",
        types,
    );
    assert_eq!(out, vec!["111"]);
}

#[test]
fn interface_polymorphism_across_two_implementations() {
    let types = r#"
        interface Op { int run(int n); }
        static class AddOne implements Op { public int run(int n) { return n + 1; } }
        static class TimesTwo implements Op { public int run(int n) { return n * 2; } }
    "#;
    let out = run_in_main(
        "Op a = new AddOne(); Op b = new TimesTwo(); System.out.println(a.run(4)); System.out.println(b.run(4));",
        types,
    );
    assert_eq!(out, vec!["5", "8"]);
}

#[test]
fn default_method_overridden_in_middle_of_hierarchy() {
    let types = r#"
        interface Speak { default String say() { return "iface"; } }
        static class BaseSpeaker implements Speak {}
        static class LoudSpeaker extends BaseSpeaker {
            public String say() { return "LOUD"; }
        }
    "#;
    let out = run_in_main(
        "Speak s = new LoudSpeaker(); System.out.println(s.say());",
        types,
    );
    assert_eq!(out, vec!["LOUD"]);
}

#[test]
fn interface_with_boolean_method() {
    let types = r#"
        interface Check { boolean ok(int n); }
        static class PositiveCheck implements Check {
            public boolean ok(int n) { return n > 0; }
        }
    "#;
    let out = run_in_main(
        "Check c = new PositiveCheck(); System.out.println(c.ok(3)); System.out.println(c.ok(-1));",
        types,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn interface_default_returns_computed_string() {
    let types = r#"
        interface Named { String name(); default String greeting() { return "hi " + name(); } }
        static class User implements Named {
            String n;
            User(String n) { this.n = n; }
            public String name() { return n; }
        }
    "#;
    let out = run_in_main(
        "Named u = new User(\"ann\"); System.out.println(u.greeting());",
        types,
    );
    assert_eq!(out, vec!["hi ann"]);
}

#[test]
fn static_interface_method_returns_boolean_flag() {
    let types = r#"
        interface Flags { static boolean enabled() { return true; } }
    "#;
    let out = run_in_main("System.out.println(Flags.enabled());", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_implements_interface_with_void_method() {
    let types = r#"
        interface Sink { void accept(int n); }
        static class PrintSink implements Sink {
            public void accept(int n) { System.out.println(n); }
        }
    "#;
    let out = run_in_main("Sink s = new PrintSink(); s.accept(42);", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn interface_extends_pair_and_class_implements_child() {
    let types = r#"
        interface A { default String a() { return "A"; } }
        interface B { default String b() { return "B"; } }
        interface AB extends A, B {}
        static class All implements AB {}
    "#;
    let out = run_in_main(
        "AB ab = new All(); System.out.println(ab.a() + ab.b());",
        types,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn default_method_chain_calls_two_defaults() {
    let types = r#"
        interface Part {
            default String partA() { return "a"; }
            default String partB() { return "b"; }
            default String whole() { return partA() + partB(); }
        }
        static class WholePart implements Part {}
    "#;
    let out = run_in_main(
        "Part p = new WholePart(); System.out.println(p.whole());",
        types,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn interface_reference_stores_concrete_implementation() {
    let types = r#"
        interface Storage { int size(); }
        static class MemoryStorage implements Storage {
            int n;
            MemoryStorage(int n) { this.n = n; }
            public int size() { return n; }
        }
    "#;
    let out = run_in_main(
        "Storage s = new MemoryStorage(5); System.out.println(s.size());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn two_classes_implement_same_interface_differently() {
    let types = r#"
        interface Sign { String mark(); }
        static class Plus implements Sign { public String mark() { return "+"; } }
        static class Minus implements Sign { public String mark() { return "-"; } }
    "#;
    let out = run_in_main(
        "Sign p = new Plus(); Sign m = new Minus(); System.out.println(p.mark()); System.out.println(m.mark());",
        types,
    );
    assert_eq!(out, vec!["+", "-"]);
}

#[test]
fn interface_static_method_accepts_two_arguments() {
    let types = r#"
        interface Math2 { static int add(int a, int b) { return a + b; } }
    "#;
    let out = run_in_main("System.out.println(Math2.add(8, 5));", types);
    assert_eq!(out, vec!["13"]);
}

#[test]
fn implementing_class_may_add_extra_public_methods() {
    let types = r#"
        interface Core { int core(); }
        static class ExtendedCore implements Core {
            public int core() { return 1; }
            public int extra() { return 9; }
        }
    "#;
    let out = run_in_main(
        "ExtendedCore e = new ExtendedCore(); System.out.println(e.core() + e.extra());",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn default_method_can_read_implementation_field_via_method() {
    let types = r#"
        interface HasLen { int length(); default boolean empty() { return length() == 0; } }
        static class Text implements HasLen {
            String s;
            Text(String s) { this.s = s; }
            public int length() { return s.length(); }
        }
    "#;
    let out = run_in_main(
        "HasLen t = new Text(\"\"); System.out.println(t.empty());",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn interface_with_multiple_static_methods() {
    let types = r#"
        interface Range {
            static int min(int a, int b) { return a < b ? a : b; }
            static int max(int a, int b) { return a > b ? a : b; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Range.min(3, 7)); System.out.println(Range.max(3, 7));",
        types,
    );
    assert_eq!(out, vec!["3", "7"]);
}

#[test]
fn class_implements_interface_and_uses_default_in_constructor_path() {
    let types = r#"
        interface Seed { default int seed() { return 4; } }
        static class Plant implements Seed {
            int grow() { return seed() * 2; }
        }
    "#;
    let out = run_in_main(
        "Plant p = new Plant(); System.out.println(p.grow());",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

use crate::helpers::compile_ok_check;

// ── Default-method diamond must be resolved explicitly ───────────────
//
// Java: when a class inherits the SAME default method from two unrelated
// interfaces it must override it (typically delegating via `X.super.m()`), or
// compilation fails. A silent pick is wrong. `X.super.m()` itself is covered
// elsewhere; this pins the UNRESOLVED diamond.
// See flexclassplan.md §4c (`AugmentationConflict::RequireExplicit`).

#[test]
fn unresolved_default_method_diamond_is_rejected() {
    assert!(!compile_ok_check(
        "interface A { default String who() { return \"a\"; } } \
         interface B { default String who() { return \"b\"; } } \
         class C implements A, B { } \
         public class Main { public static void main(String[] a) { System.out.println(new C().who()); } }",
    ));
}

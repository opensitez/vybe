use crate::helpers::run_in_main;

#[test]
fn record_components_assigned_via_constructor() {
    let types = r#"
        static record Point(int x, int y) {}
    "#;
    let out = run_in_main(
        "Point p = new Point(3, 4); System.out.println(p.x()); System.out.println(p.y());",
        types,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn record_accessor_reads_x_component() {
    let types = r#"
        static record Pair(int left, int right) {}
    "#;
    let out = run_in_main(
        "Pair p = new Pair(10, 20); System.out.println(p.left());",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn record_accessor_reads_y_component() {
    let types = r#"
        static record Pair(int left, int right) {}
    "#;
    let out = run_in_main(
        "Pair p = new Pair(10, 20); System.out.println(p.right());",
        types,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn record_equals_same_components_true() {
    let types = r#"
        static record Point(int x, int y) {}
    "#;
    let out = run_in_main(
        "Point a = new Point(1, 2); Point b = new Point(1, 2); System.out.println(a.equals(b));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn record_equals_different_components_false() {
    let types = r#"
        static record Point(int x, int y) {}
    "#;
    let out = run_in_main(
        "Point a = new Point(1, 2); Point b = new Point(1, 3); System.out.println(a.equals(b));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn record_hashcode_equal_for_equal_objects() {
    let types = r#"
        static record Point(int x, int y) {}
    "#;
    let out = run_in_main(
        "Point a = new Point(2, 3); Point b = new Point(2, 3); System.out.println(a.hashCode() == b.hashCode());",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn record_to_string_contains_component_names() {
    let types = r#"
        static record Point(int x, int y) {}
    "#;
    let out = run_in_main(
        "Point p = new Point(1, 2); System.out.println(p.toString());",
        types,
    );
    assert_eq!(out, vec!["Point[x=1, y=2]"]);
}

#[test]
fn record_compact_constructor_normalizes_negative_age() {
    let types = r#"
        static record Person(String name, int age) {
            Person {
                if (age < 0) { age = 0; }
            }
        }
    "#;
    let out = run_in_main(
        "Person p = new Person(\"anon\", -5); System.out.println(p.age());",
        types,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn record_static_field_independent_of_instances() {
    let types = r#"
        static record Counter(int id) {
            static int total = 0;
            Counter { total++; }
        }
    "#;
    let out = run_in_main(
        "Counter a = new Counter(1); Counter b = new Counter(2); System.out.println(Counter.total);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn record_custom_instance_method() {
    let types = r#"
        static record Person(String name, int age) {
            String greeting() { return "Hi " + name; }
        }
    "#;
    let out = run_in_main(
        "Person p = new Person(\"Ada\", 30); System.out.println(p.greeting());",
        types,
    );
    assert_eq!(out, vec!["Hi Ada"]);
}

#[test]
fn record_with_single_component() {
    let types = r#"
        static record Id(int value) {}
    "#;
    let out = run_in_main("Id id = new Id(99); System.out.println(id.value());", types);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn record_with_three_components() {
    let types = r#"
        static record Triple(int a, int b, int c) {}
    "#;
    let out = run_in_main(
        "Triple t = new Triple(1, 2, 3); System.out.println(t.a() + t.b() + t.c());",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn record_equality_reflexive() {
    let types = r#"
        static record Box(int size) {}
    "#;
    let out = run_in_main(
        "Box b = new Box(4); System.out.println(b.equals(b));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn record_inequality_when_one_component_differs() {
    let types = r#"
        static record Tag(String text) {}
    "#;
    let out = run_in_main(
        "Tag a = new Tag(\"a\"); Tag b = new Tag(\"b\"); System.out.println(a.equals(b));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn record_to_string_multiple_fields() {
    let types = r#"
        static record User(String name, int id) {}
    "#;
    let out = run_in_main(
        "User u = new User(\"bob\", 7); System.out.println(u.toString());",
        types,
    );
    assert_eq!(out, vec!["User[name=bob, id=7]"]);
}

#[test]
fn record_hashcode_differs_for_unequal_records() {
    let types = r#"
        static record Pair(int x, int y) {}
    "#;
    let out = run_in_main(
        "Pair a = new Pair(1, 2); Pair b = new Pair(2, 1); System.out.println(a.hashCode() == b.hashCode());",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn record_compact_constructor_trims_string_name() {
    let types = r#"
        static record Label(String text) {
            Label {
                text = text.trim();
            }
        }
    "#;
    let out = run_in_main(
        "Label l = new Label(\"  ok  \"); System.out.println(l.text());",
        types,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn record_static_method_on_record() {
    let types = r#"
        static record Version(int major, int minor) {
            static String label(int major, int minor) { return major + "." + minor; }
        }
    "#;
    let out = run_in_main("System.out.println(Version.label(1, 2));", types);
    assert_eq!(out, vec!["1.2"]);
}

#[test]
fn record_nested_in_outer_class() {
    let types = r#"
        static class Outer {
            static record Inner(int code) {}
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(5); System.out.println(inner.code());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn record_implements_interface_method() {
    let types = r#"
        static interface Named { String name(); }
        static record Person(String name, int age) implements Named {}
    "#;
    let out = run_in_main(
        "Named n = new Person(\"Ann\", 20); System.out.println(n.name());",
        types,
    );
    assert_eq!(out, vec!["Ann"]);
}

#[test]
fn record_component_reference_in_method() {
    let types = r#"
        static record Rect(int w, int h) {
            int area() { return w * h; }
        }
    "#;
    let out = run_in_main(
        "Rect r = new Rect(3, 4); System.out.println(r.area());",
        types,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn record_copy_via_constructor() {
    let types = r#"
        static record Point(int x, int y) {}
    "#;
    let out = run_in_main(
        "Point a = new Point(1, 2); Point b = new Point(a.x(), a.y()); System.out.println(b.equals(a));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn record_with_boolean_component() {
    let types = r#"
        static record Flag(boolean on) {}
    "#;
    let out = run_in_main(
        "Flag f = new Flag(true); System.out.println(f.on());",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn record_with_long_component() {
    let types = r#"
        static record Wide(long value) {}
    "#;
    let out = run_in_main(
        "Wide w = new Wide(1000L); System.out.println(w.value());",
        types,
    );
    assert_eq!(out, vec!["1000"]);
}

#[test]
fn record_equals_null_returns_false() {
    let types = r#"
        static record Unit(int n) {}
    "#;
    let out = run_in_main(
        "Unit u = new Unit(1); System.out.println(u.equals(null));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn record_equals_different_type_returns_false() {
    let types = r#"
        static record Unit(int n) {}
        static class Box { int n; Box(int n) { this.n = n; } }
    "#;
    let out = run_in_main(
        "Unit u = new Unit(1); Box b = new Box(1); System.out.println(u.equals(b));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn record_to_string_empty_name_component() {
    let types = r#"
        static record Label(String text) {}
    "#;
    let out = run_in_main(
        "Label l = new Label(\"\"); System.out.println(l.toString());",
        types,
    );
    assert_eq!(out, vec!["Label[text=]"]);
}

#[test]
fn record_static_counter_increments_per_instance() {
    let types = r#"
        static record Ticket(int id) {
            static int issued = 0;
            Ticket { issued++; }
        }
    "#;
    let out = run_in_main(
        "new Ticket(1); new Ticket(2); System.out.println(Ticket.issued);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn record_canonical_constructor_explicit() {
    let types = r#"
        static record Pair(int a, int b) {
            Pair(int a, int b) { this.a = a; this.b = b; }
        }
    "#;
    let out = run_in_main(
        "Pair p = new Pair(2, 3); System.out.println(p.a()); System.out.println(p.b());",
        types,
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn record_hashcode_stable_on_repeated_calls() {
    let types = r#"
        static record Seed(int n) {}
    "#;
    let out = run_in_main(
        "Seed s = new Seed(7); int h1 = s.hashCode(); int h2 = s.hashCode(); System.out.println(h1 == h2);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn record_two_instances_same_values_equal() {
    let types = r#"
        static record Coord(int x, int y) {}
    "#;
    let out = run_in_main(
        "Coord a = new Coord(5, 6); Coord b = new Coord(5, 6); System.out.println(a.equals(b));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn record_compact_constructor_uppercase_string() {
    let types = r#"
        static record Code(String value) {
            Code { value = value.toUpperCase(); }
        }
    "#;
    let out = run_in_main(
        "Code c = new Code(\"abc\"); System.out.println(c.value());",
        types,
    );
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn record_local_in_method() {
    let types = r#"
        static class Maker {
            int make() {
                record Local(int v) {}
                return new Local(8).v();
            }
        }
    "#;
    let out = run_in_main(
        "Maker m = new Maker(); System.out.println(m.make());",
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn record_component_order_matters_in_equals() {
    let types = r#"
        static record Pair(int first, int second) {}
    "#;
    let out = run_in_main(
        "Pair a = new Pair(1, 2); Pair b = new Pair(2, 1); System.out.println(a.equals(b));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn record_string_component_preserved() {
    let types = r#"
        static record Message(String text) {}
    "#;
    let out = run_in_main(
        "Message m = new Message(\"hello\"); System.out.println(m.text());",
        types,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn record_with_zero_components_count() {
    let types = r#"
        static record Empty() {
            int constant() { return 1; }
        }
    "#;
    let out = run_in_main(
        "Empty e = new Empty(); System.out.println(e.constant());",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn record_static_field_readable_before_instances() {
    let types = r#"
        static record Node(int id) {
            static int maxId = 100;
        }
    "#;
    let out = run_in_main("System.out.println(Node.maxId);", types);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn record_compact_constructor_clamps_upper_bound() {
    let types = r#"
        static record Score(int value) {
            Score {
                if (value > 100) { value = 100; }
            }
        }
    "#;
    let out = run_in_main(
        "Score s = new Score(150); System.out.println(s.value());",
        types,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn record_to_string_single_component() {
    let types = r#"
        static record Id(int value) {}
    "#;
    let out = run_in_main(
        "Id id = new Id(42); System.out.println(id.toString());",
        types,
    );
    assert_eq!(out, vec!["Id[value=42]"]);
}

#[test]
fn record_multiple_static_fields() {
    let types = r#"
        static record Config(String env) {
            static int retries = 3;
            static boolean debug = true;
        }
    "#;
    let out = run_in_main(
        "System.out.println(Config.retries); System.out.println(Config.debug);",
        types,
    );
    assert_eq!(out, vec!["3", "true"]);
}

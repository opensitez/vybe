use crate::helpers::{run_in_main, run_main};

#[test]
fn default_constructor_leaves_numeric_field_at_zero() {
    let types = r#"
        static class Cell {
            int value;
        }
    "#;
    let out = run_in_main("Cell c = new Cell(); System.out.println(c.value);", types);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn parameterized_constructor_sets_both_coordinates() {
    let types = r#"
        static class Point {
            int x;
            int y;
            Point(int x, int y) { this.x = x; this.y = y; }
        }
    "#;
    let out = run_in_main(
        "Point p = new Point(3, 4); System.out.println(p.x); System.out.println(p.y);",
        types,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn field_default_initialization_applies_before_constructor_body() {
    let types = r#"
        static class Counter {
            int count = 5;
            Counter() {}
        }
    "#;
    let out = run_in_main("Counter c = new Counter(); System.out.println(c.count);", types);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn field_explicit_initializer_overrides_default_zero() {
    let types = r#"
        static class Seed {
            int n = 99;
        }
    "#;
    let out = run_in_main("Seed s = new Seed(); System.out.println(s.n);", types);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn this_disambiguates_field_from_constructor_parameter() {
    let types = r#"
        static class Holder {
            int value;
            Holder(int value) { this.value = value; }
        }
    "#;
    let out = run_in_main("Holder h = new Holder(42); System.out.println(h.value);", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn this_constructor_chaining_delegates_to_overloaded_ctor() {
    let types = r#"
        static class Rect {
            int w;
            int h;
            Rect(int side) { this(side, side); }
            Rect(int w, int h) { this.w = w; this.h = h; }
        }
    "#;
    let out = run_in_main(
        "Rect r = new Rect(4); System.out.println(r.w); System.out.println(r.h);",
        types,
    );
    assert_eq!(out, vec!["4", "4"]);
}

#[test]
fn static_field_defaults_to_zero_when_uninitialized() {
    let types = r#"
        static class Stats {
            static int total;
        }
    "#;
    let out = run_in_main("System.out.println(Stats.total);", types);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn static_field_explicit_initializer_visible_before_instances() {
    let types = r#"
        static class Config {
            static int port = 3000;
        }
    "#;
    let out = run_in_main("System.out.println(Config.port);", types);
    assert_eq!(out, vec!["3000"]);
}

#[test]
fn static_method_invoked_with_qualified_class_name() {
    let types = r#"
        static class Math2 {
            static int add(int a, int b) { return a + b; }
        }
    "#;
    let out = run_in_main("System.out.println(Math2.add(10, 7));", types);
    assert_eq!(out, vec!["17"]);
}

#[test]
fn new_keyword_creates_instance_with_constructor_args() {
    let types = r#"
        static class Label {
            String text;
            Label(String text) { this.text = text; }
        }
    "#;
    let out = run_in_main("Label l = new Label(\"core\"); System.out.println(l.text);", types);
    assert_eq!(out, vec!["core"]);
}

#[test]
fn multiple_instances_have_independent_instance_fields() {
    let types = r#"
        static class Slot {
            int n;
            Slot(int n) { this.n = n; }
        }
    "#;
    let out = run_in_main(
        "Slot a = new Slot(1); Slot b = new Slot(2); System.out.println(a.n); System.out.println(b.n);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn final_field_set_once_in_constructor_is_readable() {
    let types = r#"
        static class Config {
            final int port;
            Config(int port) { this.port = port; }
        }
    "#;
    let out = run_in_main("Config c = new Config(8080); System.out.println(c.port);", types);
    assert_eq!(out, vec!["8080"]);
}

#[test]
fn final_field_with_initializer_at_declaration() {
    let types = r#"
        static class Constants {
            final int max = 100;
        }
    "#;
    let out = run_in_main("Constants c = new Constants(); System.out.println(c.max);", types);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn public_field_accessible_from_main_body() {
    let types = r#"
        static class Open {
            public int value = 7;
        }
    "#;
    let out = run_in_main("Open o = new Open(); System.out.println(o.value);", types);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn private_field_read_through_public_getter() {
    let types = r#"
        static class Vault {
            private int secret = 13;
            public int getSecret() { return secret; }
        }
    "#;
    let out = run_in_main("Vault v = new Vault(); System.out.println(v.getSecret());", types);
    assert_eq!(out, vec!["13"]);
}

#[test]
fn protected_field_visible_in_subclass() {
    let types = r#"
        static class Base {
            protected int level = 2;
        }
        static class Child extends Base {
            int read() { return level; }
        }
    "#;
    let out = run_in_main("Child c = new Child(); System.out.println(c.read());", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn tostring_override_returns_custom_representation() {
    let types = r#"
        static class User {
            String name;
            User(String name) { this.name = name; }
            public String toString() { return "User(" + name + ")"; }
        }
    "#;
    let out = run_in_main(
        "User u = new User(\"ann\"); System.out.println(u.toString());",
        types,
    );
    assert_eq!(out, vec!["User(ann)"]);
}

#[test]
fn default_tostring_on_plain_object_is_not_null() {
    let types = r#"
        static class Plain {}
    "#;
    let out = run_in_main("Plain p = new Plain(); System.out.println(p.toString() != null);", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn static_field_increments_across_multiple_constructors() {
    let types = r#"
        static class Widget {
            static int created = 0;
            Widget() { created++; }
        }
    "#;
    let out = run_in_main(
        "new Widget(); new Widget(); System.out.println(Widget.created);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn instance_method_uses_this_to_return_current_field() {
    let types = r#"
        static class Echo {
            int value;
            Echo(int value) { this.value = value; }
            int self() { return this.value; }
        }
    "#;
    let out = run_in_main("Echo e = new Echo(55); System.out.println(e.self());", types);
    assert_eq!(out, vec!["55"]);
}

#[test]
fn this_resolves_shadowed_parameter_to_field_assignment() {
    let types = r#"
        static class Pair {
            int first;
            int second;
            Pair(int first, int second) {
                this.first = first;
                this.second = second;
            }
        }
    "#;
    let out = run_in_main(
        "Pair p = new Pair(1, 2); System.out.println(p.first); System.out.println(p.second);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn implicit_no_arg_constructor_generated_when_none_declared() {
    let types = r#"
        static class Empty {
            int marker = 1;
        }
    "#;
    let out = run_in_main("Empty e = new Empty(); System.out.println(e.marker);", types);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn constructor_initializes_multiple_fields_in_one_body() {
    let types = r#"
        static class RGB {
            int r;
            int g;
            int b;
            RGB(int r, int g, int b) {
                this.r = r;
                this.g = g;
                this.b = b;
            }
        }
    "#;
    let out = run_in_main(
        "RGB c = new RGB(255, 128, 0); System.out.println(c.r); System.out.println(c.g); System.out.println(c.b);",
        types,
    );
    assert_eq!(out, vec!["255", "128", "0"]);
}

#[test]
fn two_new_expressions_yield_distinct_instances() {
    let types = r#"
        static class Node {
            int id;
            Node(int id) { this.id = id; }
        }
    "#;
    let out = run_in_main(
        "Node a = new Node(1); Node b = new Node(2); System.out.println(a.id); System.out.println(b.id);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn instance_method_reads_mutable_field_after_assignment() {
    let types = r#"
        static class Bag {
            int item = 0;
            int peek() { return item; }
        }
    "#;
    let out = run_in_main("Bag b = new Bag(); b.item = 9; System.out.println(b.peek());", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn instance_method_writes_field_visible_to_later_reads() {
    let types = r#"
        static class Store {
            int value;
            void put(int v) { value = v; }
            int get() { return value; }
        }
    "#;
    let out = run_in_main("Store s = new Store(); s.put(21); System.out.println(s.get());", types);
    assert_eq!(out, vec!["21"]);
}

#[test]
fn static_final_field_acts_as_named_constant() {
    let types = r#"
        static class Limits {
            static final int MAX = 50;
        }
    "#;
    let out = run_in_main("System.out.println(Limits.MAX);", types);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn public_method_exposes_internal_computation() {
    let types = r#"
        static class Calc {
            private int base = 3;
            public int triple(int n) { return base * n; }
        }
    "#;
    let out = run_in_main("Calc c = new Calc(); System.out.println(c.triple(4));", types);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn package_private_field_accessible_within_same_outer_class() {
    let types = r#"
        static class Note {
            String text = "ok";
        }
    "#;
    let out = run_in_main("Note n = new Note(); System.out.println(n.text);", types);
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn nested_static_class_instantiated_from_main() {
    let types = r#"
        static class Outer {
            static class Inner {
                int value = 6;
            }
        }
    "#;
    let out = run_in_main(
        "Outer.Inner inner = new Outer.Inner(); System.out.println(inner.value);",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn instance_field_increment_updates_stored_value() {
    let types = r#"
        static class Tick {
            int count = 0;
            void bump() { count++; }
        }
    "#;
    let out = run_in_main(
        "Tick t = new Tick(); t.bump(); t.bump(); System.out.println(t.count);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn boolean_field_defaults_false_then_set_true() {
    let types = r#"
        static class Switch {
            boolean on;
            Switch() { on = true; }
        }
    "#;
    let out = run_in_main("Switch s = new Switch(); System.out.println(s.on);", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn string_field_holds_assigned_literal() {
    let types = r#"
        static class Title {
            String text = "untitled";
        }
    "#;
    let out = run_in_main("Title t = new Title(); System.out.println(t.text);", types);
    assert_eq!(out, vec!["untitled"]);
}

#[test]
fn double_field_stores_fractional_value() {
    let types = r#"
        static class Measure {
            double ratio = 0.5;
        }
    "#;
    let out = run_in_main("Measure m = new Measure(); System.out.println(m.ratio);", types);
    assert_eq!(out, vec!["0.5"]);
}

#[test]
fn object_field_holds_reference_to_another_instance() {
    let types = r#"
        static class Leaf {
            int id;
            Leaf(int id) { this.id = id; }
        }
        static class Branch {
            Leaf leaf;
            Branch(Leaf leaf) { this.leaf = leaf; }
        }
    "#;
    let out = run_in_main(
        "Branch b = new Branch(new Leaf(7)); System.out.println(b.leaf.id);",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn constructor_invokes_instance_helper_method() {
    let types = r#"
        static class Init {
            int value;
            Init(int seed) { value = normalize(seed); }
            int normalize(int n) { return n + 1; }
        }
    "#;
    let out = run_in_main("Init i = new Init(4); System.out.println(i.value);", types);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn overloaded_constructors_select_by_parameter_count() {
    let types = r#"
        static class Box {
            int size;
            Box() { size = 1; }
            Box(int size) { this.size = size; }
        }
    "#;
    let out = run_in_main(
        "Box a = new Box(); Box b = new Box(9); System.out.println(a.size); System.out.println(b.size);",
        types,
    );
    assert_eq!(out, vec!["1", "9"]);
}

#[test]
fn instanceof_after_new_detects_exact_runtime_type() {
    let types = r#"
        static class Animal {}
        static class Dog extends Animal {}
    "#;
    let out = run_in_main(
        "Dog d = new Dog(); System.out.println(d instanceof Dog); System.out.println(d instanceof Animal);",
        types,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn field_reassigned_after_construction_reflects_new_value() {
    let types = r#"
        static class Mutable {
            int x = 1;
        }
    "#;
    let out = run_in_main("Mutable m = new Mutable(); m.x = 8; System.out.println(m.x);", types);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn static_field_readable_before_any_instance_created() {
    let types = r#"
        static class Registry {
            static String name = "default";
        }
    "#;
    let out = run_in_main("System.out.println(Registry.name);", types);
    assert_eq!(out, vec!["default"]);
}

#[test]
fn instance_method_returns_this_reference_for_identity_check() {
    let types = r#"
        static class SelfRef {
            SelfRef getThis() { return this; }
        }
    "#;
    let out = run_in_main(
        "SelfRef s = new SelfRef(); System.out.println(s.getThis() == s);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn custom_tostring_includes_multiple_fields() {
    let types = r#"
        static class Point {
            int x;
            int y;
            Point(int x, int y) { this.x = x; this.y = y; }
            public String toString() { return "(" + x + "," + y + ")"; }
        }
    "#;
    let out = run_in_main("Point p = new Point(2, 3); System.out.println(p.toString());", types);
    assert_eq!(out, vec!["(2,3)"]);
}

#[test]
fn tostring_delegates_to_helper_method_for_formatting() {
    let types = r#"
        static class Version {
            int major;
            int minor;
            Version(int major, int minor) { this.major = major; this.minor = minor; }
            String format() { return major + "." + minor; }
            public String toString() { return "v" + format(); }
        }
    "#;
    let out = run_in_main("Version v = new Version(1, 4); System.out.println(v.toString());", types);
    assert_eq!(out, vec!["v1.4"]);
}

#[test]
fn private_static_field_accessed_only_via_static_method() {
    let types = r#"
        static class Cache {
            private static int hits = 0;
            static void record() { hits++; }
            static int getHits() { return hits; }
        }
    "#;
    let out = run_in_main(
        "Cache.record(); Cache.record(); System.out.println(Cache.getHits());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn protected_method_visible_to_subclass_override() {
    let types = r#"
        static class Base {
            protected int baseValue() { return 10; }
        }
        static class Derived extends Base {
            int total() { return baseValue() + 5; }
        }
    "#;
    let out = run_in_main("Derived d = new Derived(); System.out.println(d.total());", types);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn class_with_only_static_members_never_needs_instance() {
    let types = r#"
        static class IdGen {
            static int next = 1;
            static int allocate() { return next++; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(IdGen.allocate()); System.out.println(IdGen.allocate());",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn new_arraylist_in_main_creates_empty_collection() {
    let out = run_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); System.out.println(list.size());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn field_array_initializer_sets_first_element() {
    let types = r#"
        static class Buffer {
            int[] data = {4, 5, 6};
        }
    "#;
    let out = run_in_main("Buffer b = new Buffer(); System.out.println(b.data[0]);", types);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn static_method_updates_static_field_from_main() {
    let types = r#"
        static class Session {
            static int active = 0;
            static void open() { active++; }
        }
    "#;
    let out = run_in_main(
        "Session.open(); Session.open(); System.out.println(Session.active);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn constructor_assigns_this_fields_before_instance_use() {
    let types = r#"
        static class Account {
            String owner;
            int balance;
            Account(String owner, int balance) {
                this.owner = owner;
                this.balance = balance;
            }
        }
    "#;
    let out = run_in_main(
        "Account a = new Account(\"sam\", 100); System.out.println(a.owner); System.out.println(a.balance);",
        types,
    );
    assert_eq!(out, vec!["sam", "100"]);
}

#[test]
fn public_static_factory_method_returns_new_instance() {
    let types = r#"
        static class Token {
            int value;
            Token(int value) { this.value = value; }
            static Token of(int value) { return new Token(value); }
        }
    "#;
    let out = run_in_main("Token t = Token.of(77); System.out.println(t.value);", types);
    assert_eq!(out, vec!["77"]);
}

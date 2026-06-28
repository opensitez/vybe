use crate::helpers::{run_in_main, run_main};

#[test]
fn default_constructor_zeroes_int_field() {
    let types = r#"
        static class Cell {
            int value;
        }
    "#;
    let out = run_in_main("Cell c = new Cell(); System.out.println(c.value);", types);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn default_constructor_zeroes_boolean_field() {
    let types = r#"
        static class Flag {
            boolean on;
        }
    "#;
    let out = run_in_main("Flag f = new Flag(); System.out.println(f.on);", types);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn default_constructor_leaves_string_field_null() {
    let types = r#"
        static class Holder {
            String text;
        }
    "#;
    let out = run_in_main(
        "Holder h = new Holder(); System.out.println(h.text);",
        types,
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn parameterized_constructor_sets_single_int_field() {
    let types = r#"
        static class Slot {
            int n;
            Slot(int n) { this.n = n; }
        }
    "#;
    let out = run_in_main("Slot s = new Slot(17); System.out.println(s.n);", types);
    assert_eq!(out, vec!["17"]);
}

#[test]
fn parameterized_constructor_sets_two_coordinates() {
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
fn parameterized_constructor_accepts_string_argument() {
    let types = r#"
        static class Label {
            String text;
            Label(String text) { this.text = text; }
        }
    "#;
    let out = run_in_main(
        r#"Label l = new Label("java"); System.out.println(l.text);"#,
        types,
    );
    assert_eq!(out, vec!["java"]);
}

#[test]
fn this_constructor_chain_delegates_to_two_arg_ctor() {
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
fn this_constructor_chain_three_level_delegation() {
    let types = r#"
        static class Cube {
            int size;
            Cube() { this(1); }
            Cube(int size) { this(size, size, size); }
            Cube(int w, int h, int d) { this.size = w + h + d; }
        }
    "#;
    let out = run_in_main("Cube c = new Cube(2); System.out.println(c.size);", types);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn super_constructor_initializes_parent_field() {
    let types = r#"
        static class Base {
            int x;
            Base(int v) { x = v; }
        }
        static class Child extends Base {
            Child(int v) { super(v); }
        }
    "#;
    let out = run_in_main("Child c = new Child(42); System.out.println(c.x);", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn super_constructor_then_child_field_assignment() {
    let types = r#"
        static class Base {
            int x;
            Base(int v) { x = v; }
        }
        static class Child extends Base {
            int y;
            Child(int x, int y) { super(x); this.y = y; }
        }
    "#;
    let out = run_in_main(
        "Child c = new Child(2, 5); System.out.println(c.x); System.out.println(c.y);",
        types,
    );
    assert_eq!(out, vec!["2", "5"]);
}

#[test]
fn overloaded_constructors_select_no_arg_vs_int() {
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
fn overloaded_constructors_select_one_vs_two_int_params() {
    let types = r#"
        static class Pair {
            int sum;
            Pair(int a) { sum = a; }
            Pair(int a, int b) { sum = a + b; }
        }
    "#;
    let out = run_in_main(
        "Pair p1 = new Pair(3); Pair p2 = new Pair(3, 4); System.out.println(p1.sum); System.out.println(p2.sum);",
        types,
    );
    assert_eq!(out, vec!["3", "7"]);
}

#[test]
fn overloaded_constructors_three_distinct_signatures() {
    let types = r#"
        static class Flex {
            int tag;
            Flex() { tag = 0; }
            Flex(int n) { tag = n; }
            Flex(int a, int b) { tag = a * 10 + b; }
        }
    "#;
    let out = run_in_main(
        "Flex a = new Flex(); Flex b = new Flex(7); Flex c = new Flex(2, 3); System.out.println(a.tag); System.out.println(b.tag); System.out.println(c.tag);",
        types,
    );
    assert_eq!(out, vec!["0", "7", "23"]);
}

#[test]
fn this_disambiguates_field_from_constructor_parameter() {
    let types = r#"
        static class Holder {
            int value;
            Holder(int value) { this.value = value; }
        }
    "#;
    let out = run_in_main(
        "Holder h = new Holder(42); System.out.println(h.value);",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn field_initializer_runs_before_constructor_body() {
    let types = r#"
        static class Counter {
            int count = 5;
            Counter() {}
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); System.out.println(c.count);",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn explicit_empty_constructor_preserves_field_initializer() {
    let types = r#"
        static class Seed {
            int n = 99;
            Seed() {}
        }
    "#;
    let out = run_in_main("Seed s = new Seed(); System.out.println(s.n);", types);
    assert_eq!(out, vec!["99"]);
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
fn super_chain_across_grandparent_levels() {
    let types = r#"
        static class Grand { int g; Grand(int g) { this.g = g; } }
        static class Parent extends Grand { Parent(int g) { super(g); } }
        static class Child extends Parent { Child(int g) { super(g); } }
    "#;
    let out = run_in_main("Child c = new Child(8); System.out.println(c.g);", types);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn this_square_constructor_sets_equal_dimensions() {
    let types = r#"
        static class Square {
            int side;
            Square(int side) { this.side = side; }
        }
    "#;
    let out = run_in_main(
        "Square s = new Square(6); System.out.println(s.side);",
        types,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn constructor_assigns_multiple_this_fields_in_body() {
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
        "RGB c = new RGB(1, 2, 3); System.out.println(c.r); System.out.println(c.g); System.out.println(c.b);",
        types,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn implicit_no_arg_constructor_on_field_only_class() {
    let types = r#"
        static class Empty {
            int marker = 1;
        }
    "#;
    let out = run_in_main(
        "Empty e = new Empty(); System.out.println(e.marker);",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn constructor_accepts_expression_in_parameter() {
    let types = r#"
        static class Calc {
            int total;
            Calc(int total) { this.total = total; }
        }
    "#;
    let out = run_in_main(
        "Calc c = new Calc(3 + 4); System.out.println(c.total);",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn multiple_instances_have_independent_state_after_ctor() {
    let types = r#"
        static class Bag {
            int size;
            Bag(int size) { this.size = size; }
        }
    "#;
    let out = run_in_main(
        "Bag a = new Bag(2); Bag b = new Bag(5); System.out.println(a.size); System.out.println(b.size);",
        types,
    );
    assert_eq!(out, vec!["2", "5"]);
}

#[test]
fn this_chain_rect_preserves_width_and_height() {
    let types = r#"
        static class Rect {
            int w;
            int h;
            Rect(int w, int h) { this.w = w; this.h = h; }
            Rect(int side) { this(side, side); }
        }
    "#;
    let out = run_in_main(
        "Rect r = new Rect(3, 7); System.out.println(r.w); System.out.println(r.h);",
        types,
    );
    assert_eq!(out, vec!["3", "7"]);
}

#[test]
fn super_passes_computed_value_to_parent() {
    let types = r#"
        static class Base {
            int x;
            Base(int v) { x = v; }
        }
        static class Child extends Base {
            Child(int a, int b) { super(a + b); }
        }
    "#;
    let out = run_in_main("Child c = new Child(2, 3); System.out.println(c.x);", types);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn overloaded_selects_two_param_ctor_over_single() {
    let types = r#"
        static class Span {
            int len;
            Span(int a) { len = a; }
            Span(int a, int b) { len = a + b; }
        }
    "#;
    let out = run_in_main("Span s = new Span(4, 6); System.out.println(s.len);", types);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn default_ctor_on_class_with_only_primitive_fields() {
    let types = r#"
        static class Plain {
            int a;
            int b;
        }
    "#;
    let out = run_in_main(
        "Plain p = new Plain(); System.out.println(p.a); System.out.println(p.b);",
        types,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn constructor_initializes_array_field_reference() {
    let types = r#"
        static class Buffer {
            int[] data;
            Buffer(int n) { data = new int[] { n, n + 1 }; }
        }
    "#;
    let out = run_in_main(
        "Buffer b = new Buffer(5); System.out.println(b.data[0]); System.out.println(b.data[1]);",
        types,
    );
    assert_eq!(out, vec!["5", "6"]);
}

#[test]
fn this_call_reaches_primary_constructor_body() {
    let types = r#"
        static class Step {
            int level;
            Step() { this(0); }
            Step(int level) { this.level = level + 1; }
        }
    "#;
    let out = run_in_main("Step s = new Step(); System.out.println(s.level);", types);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn copy_constructor_pattern_via_this_fields() {
    let types = r#"
        static class Copy {
            String a;
            String b;
            Copy(String a, String b) { this.a = a; this.b = b; }
            Copy(Copy other) { this(other.a, other.b); }
        }
    "#;
    let out = run_in_main(
        r#"Copy src = new Copy("hi", "bye"); Copy dst = new Copy(src); System.out.println(dst.a); System.out.println(dst.b);"#,
        types,
    );
    assert_eq!(out, vec!["hi", "bye"]);
}

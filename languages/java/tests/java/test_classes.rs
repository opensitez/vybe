use crate::helpers::run_in_main;

#[test]
fn constructor_initializes_instance_fields() {
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
fn instance_method_reads_mutable_field() {
    let types = r#"
        static class Counter {
            int count = 0;
            void increment() { count++; }
            int get() { return count; }
        }
    "#;
    let out = run_in_main(
        "Counter c = new Counter(); c.increment(); c.increment(); System.out.println(c.get());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn static_field_shared_across_instances() {
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
fn static_method_called_without_instance() {
    let types = r#"
        static class Math2 {
            static int add(int a, int b) { return a + b; }
        }
    "#;
    let out = run_in_main("System.out.println(Math2.add(4, 5));", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn this_disambiguates_field_from_parameter() {
    let types = r#"
        static class Holder {
            int value;
            Holder(int value) { this.value = value; }
        }
    "#;
    let out = run_in_main(
        "Holder h = new Holder(99); System.out.println(h.value);",
        types,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn final_field_cannot_be_reassigned_after_init() {
    let types = r#"
        static class Config {
            final int port;
            Config(int port) { this.port = port; }
        }
    "#;
    let out = run_in_main(
        "Config c = new Config(8080); System.out.println(c.port);",
        types,
    );
    assert_eq!(out, vec!["8080"]);
}

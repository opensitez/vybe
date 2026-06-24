use crate::helpers::run_in_main;

#[test]
fn generic_box_stores_and_returns_value() {
    let types = r#"
        static class Box<T> {
            T value;
            Box(T v) { value = v; }
            T get() { return value; }
        }
    "#;
    let out = run_in_main(
        "Box<Integer> b = new Box<>(42); System.out.println(b.get());",
        types,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn generic_method_works_with_different_type_arguments() {
    let types = r#"
        static class Util {
            static <T> T identity(T x) { return x; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Util.identity(7)); System.out.println(Util.identity(\"hi\"));",
        types,
    );
    assert_eq!(out, vec!["7", "hi"]);
}

#[test]
fn bounded_type_parameter_accepts_number_subtypes() {
    let types = r#"
        static class Numbers {
            static <T extends Number> double asDouble(T n) { return n.doubleValue(); }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Numbers.asDouble(3)); System.out.println(Numbers.asDouble(2.5));",
        types,
    );
    assert_eq!(out, vec!["3.0", "2.5"]);
}

#[test]
fn wildcard_list_accepts_multiple_element_types() {
    let types = r#"
        static class Printers {
            static void printAll(java.util.List<?> items) {
                for (Object o : items) System.out.println(o);
            }
        }
    "#;
    let out = run_in_main(
        "java.util.List<String> words = java.util.Arrays.asList(\"a\", \"b\"); Printers.printAll(words);",
        types,
    );
    assert_eq!(out, vec!["a", "b"]);
}

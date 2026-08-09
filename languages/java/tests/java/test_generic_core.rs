use crate::helpers::run_in_main;
use vybe_ast::{BindingPattern, ClassMember, Statement, StmtKind};

fn java_main_body(source: &str) -> Vec<Statement> {
    let module = vybe_language_java::parse(source).expect("java parse");
    for stmt in module.body {
        let StmtKind::ClassDecl { members, .. } = stmt.kind else {
            continue;
        };
        for member in members {
            let ClassMember::Method(method) = member else {
                continue;
            };
            let StmtKind::FunctionDecl { name, body, .. } = method.kind else {
                continue;
            };
            if name == "main" {
                return body;
            }
        }
    }
    panic!("main method not found");
}

fn java_var_type_hint(body: &[Statement], name: &str) -> Option<String> {
    for stmt in body {
        let StmtKind::VarDecl { declarations, .. } = &stmt.kind else {
            continue;
        };
        for decl in declarations {
            if matches!(&decl.pattern, BindingPattern::Ident(var) if var == name) {
                return decl.type_hint.as_ref().map(ToString::to_string);
            }
        }
    }
    None
}

#[test]
fn java_generic_type_hint_preserves_type_argument_ast() {
    let body = java_main_body(
        r#"
        class Test {
            static void main(String[] args) {
                java.util.List<String> names = java.util.Arrays.asList("a");
            }
        }
        "#,
    );

    assert_eq!(
        java_var_type_hint(&body, "names").as_deref(),
        Some("java.util.List<String>")
    );
}

#[test]
fn java_nested_generic_type_hint_preserves_arguments_ast() {
    let body = java_main_body(
        r#"
        class Test {
            static void main(String[] args) {
                java.util.Map<String, java.util.List<Integer>> table = null;
            }
        }
        "#,
    );

    assert_eq!(
        java_var_type_hint(&body, "table").as_deref(),
        Some("java.util.Map<String, java.util.List<Integer>>")
    );
}

#[test]
fn generic_box_stores_integer_payload() {
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
fn generic_box_stores_string_payload() {
    let types = r#"
        static class Box<T> {
            T value;
            Box(T v) { value = v; }
            T get() { return value; }
        }
    "#;
    let out = run_in_main(
        "Box<String> b = new Box<>(\"core\"); System.out.println(b.get());",
        types,
    );
    assert_eq!(out, vec!["core"]);
}

#[test]
fn generic_identity_method_preserves_integer() {
    let types = r#"
        static class Util {
            static <T> T identity(T x) { return x; }
        }
    "#;
    let out = run_in_main("System.out.println(Util.identity(9));", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn generic_identity_method_preserves_string() {
    let types = r#"
        static class Util {
            static <T> T identity(T x) { return x; }
        }
    "#;
    let out = run_in_main("System.out.println(Util.identity(\"vybe\"));", types);
    assert_eq!(out, vec!["vybe"]);
}

#[test]
fn bounded_number_parameter_accepts_integer() {
    let types = r#"
        static class Numbers {
            static <T extends Number> double asDouble(T n) { return n.doubleValue(); }
        }
    "#;
    let out = run_in_main("System.out.println(Numbers.asDouble(4));", types);
    assert_eq!(out, vec!["4.0"]);
}

#[test]
fn bounded_number_parameter_accepts_double() {
    let types = r#"
        static class Numbers {
            static <T extends Number> double asDouble(T n) { return n.doubleValue(); }
        }
    "#;
    let out = run_in_main("System.out.println(Numbers.asDouble(1.5));", types);
    assert_eq!(out, vec!["1.5"]);
}

#[test]
fn wildcard_list_accepts_string_elements() {
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

#[test]
fn wildcard_list_accepts_integer_elements() {
    let types = r#"
        static class Printers {
            static void printAll(java.util.List<?> items) {
                for (Object o : items) System.out.println(o);
            }
        }
    "#;
    let out = run_in_main(
        "java.util.List<Integer> nums = java.util.Arrays.asList(1, 2); Printers.printAll(nums);",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn generic_pair_holds_two_type_arguments() {
    let types = r#"
        static class Pair<A, B> {
            A first;
            B second;
            Pair(A a, B b) { first = a; second = b; }
        }
    "#;
    let out = run_in_main(
        "Pair<Integer, String> p = new Pair<>(7, \"ok\"); System.out.println(p.first); System.out.println(p.second);",
        types,
    );
    assert_eq!(out, vec!["7", "ok"]);
}

#[test]
fn generic_method_returns_first_of_two_values() {
    let types = r#"
        static class Pick {
            static <T> T first(T a, T b) { return a; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Pick.first(\"left\", \"right\"));",
        types,
    );
    assert_eq!(out, vec!["left"]);
}

#[test]
fn generic_method_returns_second_of_two_values() {
    let types = r#"
        static class Pick {
            static <T> T second(T a, T b) { return b; }
        }
    "#;
    let out = run_in_main("System.out.println(Pick.second(3, 8));", types);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn diamond_operator_infers_box_type_argument() {
    let types = r#"
        static class Box<T> {
            T value;
            Box(T v) { value = v; }
        }
    "#;
    let out = run_in_main(
        "Box<Integer> b = new Box<>(11); System.out.println(b.value);",
        types,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn nested_generic_box_reads_inner_value() {
    let types = r#"
        static class Box<T> {
            T value;
            Box(T v) { value = v; }
            T get() { return value; }
        }
    "#;
    let out = run_in_main(
        "Box<Box<Integer>> outer = new Box<>(new Box<>(5)); System.out.println(outer.get().get());",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn extends_wildcard_reads_number_list_size() {
    let types = r#"
        static class Sizes {
            static int count(java.util.List<? extends Number> nums) { return nums.size(); }
        }
    "#;
    let out = run_in_main(
        "java.util.List<Integer> nums = java.util.Arrays.asList(1, 2, 3); System.out.println(Sizes.count(nums));",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn generic_static_factory_wraps_value() {
    let types = r#"
        static class Holder<T> {
            T value;
            Holder(T v) { value = v; }
            static <T> Holder<T> of(T v) { return new Holder<>(v); }
        }
    "#;
    let out = run_in_main(
        "Holder<String> h = Holder.of(\"wrap\"); System.out.println(h.value);",
        types,
    );
    assert_eq!(out, vec!["wrap"]);
}

#[test]
fn class_level_bounded_type_param_accepts_integer() {
    let types = r#"
        static class NumBox<T extends Number> {
            T n;
            NumBox(T n) { this.n = n; }
            double asDouble() { return n.doubleValue(); }
        }
    "#;
    let out = run_in_main(
        "NumBox<Integer> box = new NumBox<>(6); System.out.println(box.asDouble());",
        types,
    );
    assert_eq!(out, vec!["6.0"]);
}

#[test]
fn generic_void_method_prints_argument() {
    let types = r#"
        static class Echo {
            static <T> void show(T value) { System.out.println(value); }
        }
    "#;
    let out = run_in_main("Echo.show(\"ping\");", types);
    assert_eq!(out, vec!["ping"]);
}

#[test]
fn generic_method_pair_returns_second_value() {
    let types = r#"
        static class PairUtil {
            static <T> T second(T a, T b) { return b; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(PairUtil.second(\"first\", \"second\"));",
        types,
    );
    assert_eq!(out, vec!["second"]);
}

#[test]
fn generic_class_method_uses_type_param_in_parameter() {
    let types = r#"
        static class Bag<T> {
            T item;
            void put(T value) { item = value; }
            T peek() { return item; }
        }
    "#;
    let out = run_in_main(
        "Bag<Integer> bag = new Bag<>(); bag.put(99); System.out.println(bag.peek());",
        types,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn wildcard_parameter_accepts_mixed_object_list() {
    let types = r#"
        static class Head {
            static Object first(java.util.List<?> items) { return items.get(0); }
        }
    "#;
    let out = run_in_main(
        "java.util.List<String> list = java.util.Arrays.asList(\"x\", \"y\"); System.out.println(Head.first(list));",
        types,
    );
    assert_eq!(out, vec!["x"]);
}

#[test]
fn generic_method_max_of_two_integers() {
    let types = r#"
        static class Math2 {
            static <T extends Number> int maxInt(T a, T b) {
                int x = a.intValue();
                int y = b.intValue();
                return x >= y ? x : y;
            }
        }
    "#;
    let out = run_in_main("System.out.println(Math2.maxInt(4, 9));", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn generic_class_field_mutation_through_setter() {
    let types = r#"
        static class Cell<T> {
            T data;
            void set(T v) { data = v; }
            T get() { return data; }
        }
    "#;
    let out = run_in_main(
        "Cell<String> c = new Cell<>(); c.set(\"set\"); System.out.println(c.get());",
        types,
    );
    assert_eq!(out, vec!["set"]);
}

#[test]
fn super_wildcard_list_accepts_integer_add() {
    let types = r#"
        static class Sink {
            static void add(java.util.List<? super Integer> dest, Integer value) {
                dest.add(value);
            }
        }
    "#;
    let out = run_in_main(
        "java.util.ArrayList<Integer> list = new java.util.ArrayList<Integer>(); Sink.add(list, 4); System.out.println(list.get(0));",
        types,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn generic_method_with_explicit_class_context() {
    let types = r#"
        static class Convert {
            static <T> String asString(T value) { return "" + value; }
        }
    "#;
    let out = run_in_main("System.out.println(Convert.asString(12));", types);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn bounded_type_reads_int_value_from_number() {
    let types = r#"
        static class Read {
            static <T extends Number> int asInt(T n) { return n.intValue(); }
        }
    "#;
    let out = run_in_main("System.out.println(Read.asInt(3.9));", types);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn generic_triple_holder_exposes_three_fields() {
    let types = r#"
        static class Triple<A, B, C> {
            A a; B b; C c;
            Triple(A a, B b, C c) { this.a = a; this.b = b; this.c = c; }
        }
    "#;
    let out = run_in_main(
        "Triple<Integer, String, Boolean> t = new Triple<>(1, \"t\", true); System.out.println(t.a); System.out.println(t.b); System.out.println(t.c);",
        types,
    );
    assert_eq!(out, vec!["1", "t", "true"]);
}

#[test]
fn wildcard_prints_each_element_in_number_list() {
    let types = r#"
        static class Dump {
            static void dump(java.util.List<? extends Number> nums) {
                for (Number n : nums) System.out.println(n.intValue());
            }
        }
    "#;
    let out = run_in_main(
        "java.util.List<Integer> nums = java.util.Arrays.asList(2, 4); Dump.dump(nums);",
        types,
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn generic_method_returns_length_of_string() {
    let types = r#"
        static class Len {
            static <T extends String> int length(T s) { return s.length(); }
        }
    "#;
    let out = run_in_main("System.out.println(Len.length(\"java\"));", types);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn generic_class_constructor_sets_initial_value() {
    let types = r#"
        static class Slot<T> {
            final T value;
            Slot(T v) { value = v; }
        }
    "#;
    let out = run_in_main(
        "Slot<Double> s = new Slot<>(2.5); System.out.println(s.value);",
        types,
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn generic_identity_chain_preserves_final_string() {
    let types = r#"
        static class Chain {
            static <T> T link(T value) { return value; }
        }
    "#;
    let out = run_in_main(
        "String s = Chain.link(Chain.link(\"end\")); System.out.println(s);",
        types,
    );
    assert_eq!(out, vec!["end"]);
}

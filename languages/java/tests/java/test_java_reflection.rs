use crate::helpers::{run_in_main, run_main};

#[test]
fn class_for_name_simple_name() {
    let out = run_main(r#"System.out.println(Class.forName("java.lang.String").getSimpleName());"#);
    assert_eq!(out, vec!["String"]);
}

#[test]
fn class_literal_canonical_name() {
    let out = run_main(r#"System.out.println(String.class.getCanonicalName());"#);
    assert_eq!(out, vec!["String"]);
}

#[test]
fn class_literal_array_predicate() {
    let out = run_main(r#"System.out.println(int[].class.isArray());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_literal_array_component_type() {
    let out = run_main(r#"System.out.println(String[].class.getComponentType().getSimpleName());"#);
    assert_eq!(out, vec!["String"]);
}

#[test]
fn class_literal_primitive_predicate() {
    let out = run_main(r#"System.out.println(int.class.isPrimitive());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn interface_literal_predicate() {
    let types = r#"interface Marker { void run(); }"#;
    let out = run_in_main("System.out.println(Marker.class.isInterface());", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_literal_predicate() {
    let types = r#"enum Mode { ON, OFF }"#;
    let out = run_in_main("System.out.println(Mode.class.isEnum());", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_literal_superclass_name() {
    let types = r#"static class Base {} static class Child extends Base {}"#;
    let out = run_in_main(
        "System.out.println(Child.class.getSuperclass().getSimpleName());",
        types,
    );
    assert_eq!(out, vec!["Base"]);
}

#[test]
fn class_for_name_package_name() {
    let out =
        run_main(r#"System.out.println(Class.forName("java.util.ArrayList").getPackageName());"#);
    assert_eq!(out, vec!["java.util"]);
}

#[test]
fn class_declared_interfaces() {
    let types = r#"interface Marker {} static class Task implements Marker {}"#;
    let out = run_in_main(
        "System.out.println(Task.class.getInterfaces().length);
         System.out.println(Task.class.getInterfaces()[0].getSimpleName());",
        types,
    );
    assert_eq!(out, vec!["1", "Marker"]);
}

#[test]
fn class_declared_nested_classes() {
    let types = r#"static class Outer { static class Inner {} }"#;
    let out = run_in_main(
        "System.out.println(Outer.class.getDeclaredClasses().length);
         System.out.println(Outer.class.getDeclaredClasses()[0].getSimpleName());",
        types,
    );
    assert_eq!(out, vec!["1", "Inner"]);
}

#[test]
fn class_modifiers_predicates() {
    let types = r#"private static class Hidden {}"#;
    let out = run_in_main(
        "System.out.println(java.lang.reflect.Modifier.isPrivate(Hidden.class.getModifiers()));
         System.out.println(java.lang.reflect.Modifier.isStatic(Hidden.class.getModifiers()));",
        types,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn class_field_modifiers_predicates() {
    let types = r#"static class Box { private final int value = 1; }"#;
    let out = run_in_main(
        r#"System.out.println(java.lang.reflect.Modifier.isPrivate(Box.class.getDeclaredField("value").getModifiers()));
           System.out.println(java.lang.reflect.Modifier.isFinal(Box.class.getDeclaredField("value").getModifiers()));"#,
        types,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn class_method_modifiers_predicates() {
    let types = r#"static class Ops { public static final int run() { return 1; } }"#;
    let out = run_in_main(
        r#"System.out.println(java.lang.reflect.Modifier.isPublic(Ops.class.getDeclaredMethod("run").getModifiers()));
           System.out.println(java.lang.reflect.Modifier.isStatic(Ops.class.getDeclaredMethod("run").getModifiers()));
           System.out.println(java.lang.reflect.Modifier.isFinal(Ops.class.getDeclaredMethod("run").getModifiers()));"#,
        types,
    );
    assert_eq!(out, vec!["true", "true", "true"]);
}

#[test]
fn class_constructor_modifiers_predicates() {
    let types = r#"static class Made { protected Made() {} }"#;
    let out = run_in_main(
        "System.out.println(java.lang.reflect.Modifier.isProtected(Made.class.getDeclaredConstructor().getModifiers()));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_literal_assignable_from_parent() {
    let types = r#"static class Base {} static class Child extends Base {}"#;
    let out = run_in_main(
        "System.out.println(Base.class.isAssignableFrom(Child.class));",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn class_literal_assignable_from_not_reverse() {
    let types = r#"static class Base {} static class Child extends Base {}"#;
    let out = run_in_main(
        "System.out.println(Child.class.isAssignableFrom(Base.class));",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn class_declared_fields_length() {
    let types = r#"static class Box { int value; String name; }"#;
    let out = run_in_main(
        "System.out.println(Box.class.getDeclaredFields().length);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn class_declared_field_get_name() {
    let types = r#"static class Box { int value; String name; }"#;
    let out = run_in_main(
        r#"System.out.println(Box.class.getDeclaredField("value").getName());"#,
        types,
    );
    assert_eq!(out, vec!["value"]);
}

#[test]
fn class_declared_field_get_type() {
    let types = r#"static class Box { int value; String name; }"#;
    let out = run_in_main(
        r#"System.out.println(Box.class.getDeclaredField("name").getType().getSimpleName());"#,
        types,
    );
    assert_eq!(out, vec!["String"]);
}

#[test]
fn class_declared_field_get_value() {
    let types = r#"static class Box { int value; Box(int v) { value = v; } }"#;
    let out = run_in_main(
        r#"Box b = new Box(9); System.out.println(Box.class.getDeclaredField("value").get(b));"#,
        types,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn class_declared_field_set_value() {
    let types = r#"static class Box { int value; Box(int v) { value = v; } }"#;
    let out = run_in_main(
        r#"Box b = new Box(1); Box.class.getDeclaredField("value").set(b, 7); System.out.println(b.value);"#,
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn class_declared_methods_length() {
    let types = r#"
        static class Ops {
            int inc(int x) { return x + 1; }
            int sum(int a, int b) { return a + b; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Ops.class.getDeclaredMethods().length);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn class_declared_methods_indexed_token_name() {
    let types = r#"
        static class Ops {
            int inc(int x) { return x + 1; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Ops.class.getDeclaredMethods()[0].getName());"#,
        types,
    );
    assert_eq!(out, vec!["inc"]);
}

#[test]
fn class_declared_methods_indexed_token_invoke() {
    let types = r#"
        static class Ops {
            int inc(int x) { return x + 1; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Ops.class.getDeclaredMethods()[0].invoke(new Ops(), 4));"#,
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn class_declared_method_parameter_count() {
    let types = r#"
        static class Ops {
            int inc(int x) { return x + 1; }
            int sum(int a, int b) { return a + b; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Ops.class.getDeclaredMethod("sum", int.class, int.class).getParameterCount());"#,
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn class_declared_method_name_and_return_type() {
    let types = r#"
        static class Ops {
            int inc(int x) { return x + 1; }
            String label() { return "ok"; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Ops.class.getDeclaredMethod("label").getName());
           System.out.println(Ops.class.getDeclaredMethod("label").getReturnType().getSimpleName());"#,
        types,
    );
    assert_eq!(out, vec!["label", "String"]);
}

#[test]
fn class_declared_method_parameter_types_length() {
    let types = r#"
        static class Ops {
            int sum(int a, int b) { return a + b; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Ops.class.getDeclaredMethod("sum", int.class, int.class).getParameterTypes().length);"#,
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn class_declared_method_invoke_instance() {
    let types = r#"
        static class Ops {
            int sum(int a, int b) { return a + b; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Ops.class.getDeclaredMethod("sum", int.class, int.class).invoke(new Ops(), 2, 3));"#,
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn class_declared_method_invoke_static() {
    let types = r#"
        static class Ops {
            static int twice(int x) { return x * 2; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Ops.class.getDeclaredMethod("twice", int.class).invoke(null, 4));"#,
        types,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn class_declared_method_invoke_object_array_args() {
    let types = r#"
        static class Ops {
            String join(String a, String b) { return a + b; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Ops.class.getDeclaredMethod("join", String.class, String.class).invoke(new Ops(), new Object[]{"a", "b"}));"#,
        types,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn class_declared_constructors_length() {
    let types = r#"
        static class Made {
            Made() {}
            Made(int x) {}
        }
    "#;
    let out = run_in_main(
        "System.out.println(Made.class.getDeclaredConstructors().length);",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn class_declared_constructors_indexed_token_new_instance() {
    let types = r#"
        static class Made {
            int value;
            Made(int x) { value = x; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Made.class.getDeclaredConstructors()[0].newInstance(13).value);",
        types,
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn class_declared_constructor_parameter_count() {
    let types = r#"
        static class Made {
            Made() {}
            Made(int x) {}
        }
    "#;
    let out = run_in_main(
        "System.out.println(Made.class.getDeclaredConstructor(int.class).getParameterCount());",
        types,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn class_declared_constructor_name_and_parameter_types() {
    let types = r#"
        static class Made {
            Made() {}
            Made(int x) {}
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Made.class.getDeclaredConstructor(int.class).getName());
           System.out.println(Made.class.getDeclaredConstructor(int.class).getParameterTypes().length);"#,
        types,
    );
    assert_eq!(out, vec!["Made", "1"]);
}

#[test]
fn class_declared_constructor_new_instance() {
    let types = r#"
        static class Made {
            int value;
            Made(int x) { value = x; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Made.class.getDeclaredConstructor(int.class).newInstance(11).value);",
        types,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn class_declared_constructor_new_instance_object_array_args() {
    let types = r#"
        static class Made {
            String value;
            Made(String x) { value = x; }
        }
    "#;
    let out = run_in_main(
        r#"System.out.println(Made.class.getDeclaredConstructor(String.class).newInstance(new Object[]{"ok"}).value);"#,
        types,
    );
    assert_eq!(out, vec!["ok"]);
}

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
    static_nested_access,
    "System.out.println(Container.Inner.value);",
    "static class Container { static int value = 3; static class Inner { static int value = 4; } }",
    "4"
);
jm!(
    static_nested_call_method,
    "System.out.println(Container2.Inner.from(5));",
    "static class Container2 { static class Inner { static int from(int x) { return x + 1; } } }",
    "6"
);
jm!(
    member_nested_access,
    "Outer o = new Outer(); System.out.println(o.new Inner().value);",
    "static class Outer { int value = 3; class Inner { int value = 9; } }",
    "9"
);
jm!(
    inner_construction,
    "Outer2 o = new Outer2(1); System.out.println(o.new Inner().value + o.base);",
    "static class Outer2 { int base; Outer2(int base) { this.base = base; } class Inner { int value = base * 2; } }",
    "3"
);
jm!(
    inner_with_method,
    "Outer3 o = new Outer3(4); System.out.println(o.new Inner().sum(1));",
    "static class Outer3 { int base; Outer3(int base) { this.base = base; } class Inner { int sum(int x) { return x + base; } } }",
    "5"
);
jm!(
    nested_class_shadowing,
    "Container3 c = new Container3(); System.out.println(c.inner.value);",
    "static class Container3 { class inner { int value = 1; } inner inner = new inner(); }",
    "1"
);
jm!(
    static_nested_uses_outer_static,
    "System.out.println(Container4.Inner.VALUE);",
    "static class Container4 { static int base = 2; static class Inner { static int VALUE = Container4.base + 1; } }",
    "3"
);
jm!(
    nested_chain,
    "Container5 c = new Container5(); System.out.println(c.value + c.new Outer().new Inner().value);",
    "static class Container5 { int value = 1; class Outer { int value = 2; class Inner { int value = 3; } } }",
    "4"
);
jm!(
    anonymous_local_in_method,
    "System.out.println(new Factory().create().value);",
    "static class Factory { Holder create() { return new Holder(4); } static class Holder { int value; Holder(int value) { this.value = value; } } }",
    "4"
);
jm!(
    nested_generic_like,
    "Outer6 o = new Outer6(); System.out.println(o.new Inner(3).value);",
    "static class Outer6 { class Inner { int value; Inner(int v) { value = v; } } }",
    "3"
);
jm!(
    local_constant_holder,
    "System.out.println(Wrapper.Nested.constant());",
    "static class Wrapper { static class Nested { static int constant() { return 8; } } }",
    "8"
);
jm!(
    double_nested_static,
    "System.out.println(Level1.Level2.Level3.value);",
    "static class Level1 { static class Level2 { static class Level3 { static int value = 12; } } }",
    "12"
);
jm!(
    member_nested_field_read,
    "Outer7 o = new Outer7(6); System.out.println(o.new Inner().twice());",
    "static class Outer7 { int base; Outer7(int base) { this.base = base; } class Inner { int twice() { return base * 2; } } }",
    "12"
);
jm!(
    member_nested_field_update,
    "Outer8 o = new Outer8(); System.out.println(o.new Inner(2).value);",
    "static class Outer8 { int delta = 1; class Inner { int value; Inner(int v) { value = v + delta; } } }",
    "3"
);
jm!(
    local_reference_to_inner,
    "Outer9 o = new Outer9(); System.out.println(o.make().value);",
    "static class Outer9 { int base = 4; class Inner { int value; Inner(int v) { value = base + v; } } Inner make() { return new Inner(1); } }",
    "5"
);
jm!(
    two_inner_instances,
    r#"Outer10 o = new Outer10(); Outer10.Inner i = o.new Inner(); Outer10.Inner j = o.new Inner(); System.out.println(i.value + "," + j.value);"#,
    "static class Outer10 { class Inner { int value = 1; } }",
    "1,1"
);
jm!(
    nested_builder,
    "Outer11 o = new Outer11(); System.out.println(o.new Builder().build(2));",
    "static class Outer11 { class Builder { int build(int v) { return v + 1; } } }",
    "3"
);
jm!(
    static_nested_return_type,
    "System.out.println(Container7.Inner.make(4));",
    "static class Container7 { static class Inner { static int make(int x) { return x * 2; } } }",
    "8"
);
jm!(
    member_nested_static_value,
    "Outer11 b = new Outer11(3); System.out.println(b.new Holder().from());",
    "static class Outer11 { int base; Outer11(int base) { this.base = base; } class Holder { int from() { return base + 1; } } }",
    "4"
);
jm!(
    nested_array_creation,
    r#"Outer12 o = new Outer12(); System.out.println(o.new Inner(1).value + "," + o.new Inner(2).value);"#,
    "static class Outer12 { class Inner { int value; Inner(int v) { value = v; } } }",
    "1,2"
);
jm!(
    nested_private_access,
    "Outer13 o = new Outer13(); System.out.println(o.new Inner().value);",
    "static class Outer13 { private int base = 2; class Inner { int value = base + 1; } }",
    "3"
);
jm!(
    outer_method_from_inner,
    "Outer14 o = new Outer14(); System.out.println(o.new Inner().value());",
    "static class Outer14 { int base = 5; class Inner { int value() { return Outer14.this.base + 1; } } }",
    "6"
);
jm!(
    nested_method_chain,
    "Outer15 o = new Outer15(); System.out.println(o.new Inner().next().value);",
    "static class Outer15 { int base = 2; class Inner { Inner next() { return this; } int value = base; } }",
    "2"
);
jm!(
    static_nested_from_main,
    "System.out.println(Factory2.Inner.ONE);",
    "static class Factory2 { static class Inner { static int ONE = 1; } }",
    "1"
);
jm!(
    nested_class_constructor_args,
    "Outer16 o = new Outer16(); System.out.println(o.new Inner(4).value);",
    "static class Outer16 { class Inner { int value; Inner(int value) { this.value = value; } } }",
    "4"
);
jm!(
    member_inner_from_static_context,
    "Outer17 o = new Outer17(); System.out.println(o.valueOf(3));",
    "static class Outer17 { int value = 1; int valueOf(int v) { return new Inner(v).value; } class Inner { int value; Inner(int v) { value = Outer17.this.value + v; } } }",
    "4"
);

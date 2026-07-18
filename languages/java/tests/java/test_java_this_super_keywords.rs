use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(this_ref_in_instance, "System.out.println(new Box(3).value);", "static class Box { int value; Box(int value) { this.value = value; } int value() { return this.value; } }", "3");
jm!(this_in_method_chain, "System.out.println(new Fluent().id().id().value);", "static class Fluent { int value = 1; Fluent id() { value++; return this; } }", "3");
jm!(constructor_this_simple, "System.out.println(new Packet().value);", "static class Packet { int value; Packet() { this(5); } Packet(int value) { this.value = value; } }", "5");
jm!(constructor_this_chain, "System.out.println(new Packet2().value);", "static class Packet2 { int value; Packet2() { this(2); } Packet2(int value) { this.value = value + 1; } Packet2(int value, int inc) { this(value + inc); } }", "3");
jm!(this_access_field, "System.out.println(new Holder(4).sum(3));", "static class Holder { int value; Holder(int value) { this.value = value; } int sum(int d) { return this.value + d; } }", "7");
jm!(this_for_nested_call, "System.out.println(new Chain(4).next().value);", "static class Chain { int value; Chain(int value) { this.value = value; } Chain next() { return this; } }", "4");
jm!(this_in_static_context_not_allowed, "System.out.println(new NotStatic().ok());", "static class NotStatic { int value = 1; int ok() { return this.value; } }", "1");
jm!(super_field_access, "System.out.println(new Child().name);", "static class Parent { String name = \"parent\"; } static class Child extends Parent { String name = \"child\"; String readParent() { return super.name; } String readThis() { return this.name; } }", "child");
jm!(super_method_access, "System.out.println(new Sub(2).add());", "static class Base { int base = 1; int add() { return base; } } static class Sub extends Base { int value; Sub(int value) { this.value = value; } int add() { return super.add() + value; } }", "3");
jm!(super_constructor_chain, "System.out.println(new Green().value);", "static class Root { int value; Root(int value) { this.value = value; } } static class Green extends Root { Green() { super(7); } }", "7");
jm!(super_constructor_default, "System.out.println(new Blue().label);", "static class Root2 { String label = \"r\"; Root2() { label = \"root\"; } Root2(String x) { label = x; } } static class Blue extends Root2 { String label = \"blue\"; Blue() { super(); } }", "blue");
jm!(this_from_default_ctor, "System.out.println(new D().value);", "static class D { int value; D() { this(1); } D(int v) { this.value = v; } }", "1");
jm!(this_return_current, "System.out.println(new R(2).next().value);", "static class R { int value; R(int value) { this.value = value; } R next() { return this; } }", "2");
jm!(this_in_private_method, "System.out.println(new N(4).doubleValue());", "static class N { int value; N(int value) { this.value = value; } private int value() { return value; } int doubleValue() { return this.value() * 2; } }", "8");
jm!(super_this_mix, "System.out.println(new S().value);", "static class A { int base = 1; S(){} } static class S extends A { int value; S() { this(2); } S(int v) { this.value = super.base + v; } }", "3");
jm!(chained_constructor_with_defaults, "System.out.println(new T().value);", "static class T { int value; T() { this(1); } T(int a) { this(a, 2); } T(int a, int b) { value = a + b; } }", "3");
jm!(this_in_boolean_method, "System.out.println(new U().is(2));", "static class U { int value = 2; boolean is(int x) { return this.value == x; } }", "true");
jm!(this_from_array_init, "System.out.println(new ArrBox(2).total());", "static class ArrBox { int value; ArrBox(int v) { this.value = v; } ArrBox total(int times) { int r = 0; for (int i = 0; i < times; i++) r += this.value; return r; } ArrBox total() { return this; } int total() { return this.value; } }", "2");
jm!(this_passed_as_arg, "System.out.println(new Caller().call().value);", "static class Caller { int value = 5; Caller call() { return this; } }", "5");
jm!(this_in_equals_chain, "System.out.println(new Eq(1).same(1) ? \"y\" : \"n\");", "static class Eq { int value; Eq(int v) { this.value = v; } boolean same(int x) { return this.value == x; } }", "y");
jm!(super_interface_default, "System.out.println(new D2().label());", "interface Named { default String label() { return \"base\"; } } static class D2 implements Named { public String label() { return super.label() + \"2\"; } }", "base2");
jm!(this_and_super_on_methods, "System.out.println(new C2(2).score());", "static class Base2 { int add(int a) { return a + 1; } } static class C2 extends Base2 { int value; C2(int value) { this.value = value; } int score() { return super.add(this.value); } }", "3");
jm!(this_in_toString_override, "System.out.println(new Name(\"x\").label());", "static class Name { String v; Name(String v) { this.v = v; } String label() { return this.v; } }", "x");
jm!(this_multiple_updates, "System.out.println(new V().set(1).set(2).value);", "static class V { int value = 0; V set(int v) { this.value = v; return this; } }", "2");
jm!(this_in_nested_object, "System.out.println(new Outer().inner().value);", "static class Outer { int value = 9; Inner inner() { return new Inner(); } class Inner { int value = Outer.this.value; } }", "9");
jm!(this_on_static_class, "System.out.println(new Holder2(5).value);", "static class Holder2 { int value; Holder2(int value) { this.value = value; } static class Builder { Holder2 build(int v) { return new Holder2(v); } } }", "5");

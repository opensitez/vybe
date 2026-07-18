use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(subclass_inherits_parent_field, "System.out.println(new Child().value);", "static class Parent { int value = 1; } static class Child extends Parent {}", "1");
jm!(subclass_override_simple, "System.out.println(new Child().name());", "static class Parent { String name() { return \"parent\"; } } static class Child extends Parent { String name() { return \"child\"; } }", "child");
jm!(polymorphism_calls_override, "Base b = new Child(); System.out.println(b.name());", "static class Base { String name() { return \"base\"; } } static class Child extends Base { String name() { return \"child\"; } }", "child");
jm!(super_implementation_from_child, "System.out.println(new Child().name());", "static class Base { String name() { return \"base\"; } } static class Child extends Base { String name() { return super.name() + \"+child\"; } }", "base+child");
jm!(constructor_chain, "System.out.println(new Child().n);", "static class Base { int n; Base(){n = 1;} } static class Child extends Base { Child(){ super(); n += 2; } }", "3");
jm!(array_of_polymorphic_types, "Base[] arr = {new Base(), new Child()}; System.out.println(arr[1].name());", "static class Base { String name() { return \"base\"; } } static class Child extends Base { String name() { return \"child\"; } }", "child");
jm!(constructor_overload_with_super, "System.out.println(new Child(4).value());", "static class Parent { int base; Parent(){ this(1); } Parent(int v){ base=v; } int value(){ return base; } } static class Child extends Parent { Child(int v){ super(v); } }", "4");
jm!(abstract_class_dispatch, "System.out.println(new Square().area());", "abstract class Shape { abstract int area(); } static class Square extends Shape { int area() { return 9; } }", "9");
jm!(interface_default_uses_implementation, "System.out.println(new Implementor().name());", "interface Named { default String name() { return \"named\"; } } static class Implementor implements Named {}", "named");
jm!(interface_default_override, "System.out.println(new Loud().name());", "interface Named { default String name() { return \"named\"; } } static class Loud implements Named { public String name() { return \"LOUD\"; } }", "LOUD");
jm!(interface_static_call, "System.out.println(Factory.value());", "interface Factory { static String value() { return \"v1\"; } }", "v1");
jm!(multiple_interface_default, "System.out.println(new Mix().tag());", "interface A { default String tag() { return \"A\"; } } interface B { default String tag() { return \"B\"; } } static class Mix implements A, B { public String tag() { return A.super.tag() + B.super.tag(); } }", "AB");
jm!(abstract_through_reference, "System.out.println(new Dog().sound());", "abstract class Animal { abstract String sound(); } static class Dog extends Animal { String sound() { return \"woof\"; } }", "woof");
jm!(interface_dispatch_on_reference, "Talker t = new Cat(); System.out.println(t.says());", "interface Talker { String says(); } static class Cat implements Talker { public String says() { return \"meow\"; } }", "meow");
jm!(instanceof_true_path, "System.out.println(new Tiger() instanceof Animal);", "abstract class Animal {} static class Tiger extends Animal {}", "true");
jm!(instanceof_false_path, "System.out.println(new Animal() instanceof Tiger);", "static class Animal {} static class Tiger extends Animal {}", "false");
jm!(safe_downcast, "Base b = new Child(); Child c = (Child)b; System.out.println(c.ok());", "static class Base {} static class Child extends Base { int ok() { return 12; } }", "12");
jm!(failed_downcast_with_catch, "Base b = new Base(); try { Child c = (Child)b; System.out.println(c.ok()); } catch (ClassCastException e) { System.out.println(\"bad\"); }", "static class Base {} static class Child extends Base { int ok() { return 1; } }", "bad");
jm!(final_class_cannot_extend, "Dog d = new Dog(); System.out.println(d.sound());", "static final class Dog { String sound() { return \"woof\"; } }", "woof");
jm!(shadowed_field, "Parent p = new Parent(); Child c = new Child(); System.out.println(c.value + \"/\" + p.value);", "static class Parent { int value = 1; } static class Child extends Parent { int value = 2; }", "2/1");
jm!(to_string_override_chain, "Base b = new Child(); System.out.println(b.toString());", "static class Base { public String toString() { return \"base\"; } } static class Child extends Base { public String toString() { return \"child\"; } }", "child");
jm!(get_class_name_runtime, "Base b = new Child(); System.out.println(b.getClass().getSimpleName());", "static class Base {} static class Child extends Base {}", "Child");
jm!(sealed_style_hierarchy, "Shape s = new Square(4); System.out.println(s.area());", "static class Shape { int area() { return 0; } } static class Square extends Shape { int side; Square(int side) { this.side = side; } int area() { return side * side; } }", "16");
jm!(deep_override_chain, "Leaf l = new Leaf(); System.out.println(l.label());", "abstract class Root { String label() { return \"root\"; } } static class Mid extends Root { String label() { return \"mid\"; } } static class Leaf extends Mid { String label() { return \"leaf\"; } }", "leaf");
jm!(hierarchy_instanceof_checks, "Animal a = new Dog(); System.out.println(a instanceof Dog);", "static class Animal {} static class Dog extends Animal {}", "true");
jm!(method_from_interfaces_and_classes, "System.out.println(new SmartDog().describe());", "interface Walker { default String role() { return \"walk\"; } } static class Dog { String bark() { return \"bark\"; } } static class SmartDog extends Dog implements Walker { String describe() { return bark() + \"/\" + role(); } }", "bark/walk");
jm!(diamond_conflict_resolved, "System.out.println(new Combined().value());", "interface A { default int value() { return 1; } } interface B { default int value() { return 2; } } static class Combined implements A, B { public int value() { return A.super.value() + B.super.value(); } }", "3");
jm!(this_calls_base, "System.out.println(new Child().token());", "static class Parent { int token() { return 1; } } static class Child extends Parent { int token() { return super.token() + 1; } }", "2");
jm!(downcast_to_interface, "Object o = new Impl(); Speaker s = (Speaker)o; System.out.println(s.say());", "interface Speaker { String say(); } static class Impl implements Speaker { public String say() { return \"ok\"; } }", "ok");

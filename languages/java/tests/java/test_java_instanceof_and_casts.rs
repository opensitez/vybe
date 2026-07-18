use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(instanceof_true, "System.out.println(new Cat() instanceof Animal);", "static class Animal {} static class Cat extends Animal {}", "true");
jm!(instanceof_false, "System.out.println(new Animal() instanceof Cat);", "static class Animal {} static class Cat extends Animal {}", "false");
jm!(instanceof_array, "System.out.println(names instanceof Object);", "String[] names = new String[]{\"a\"};", "true");
jm!(instanceof_null_false, "String s = null; System.out.println(s instanceof String);", "", "false");
jm!(cast_success, "Object x = new Integer(3); System.out.println(((Integer)x).intValue());", "", "3");
jm!(cast_with_check, "Object x = new Integer(4); int n = x instanceof Integer ? ((Integer)x).intValue() : 0; System.out.println(n);", "", "4");
jm!(cast_bad_fail, "Object x = 1; try { String s = (String)x; System.out.println(\"bad\"); } catch (ClassCastException e) { System.out.println(\"ok\"); }", "", "ok");
jm!(cast_interface_true, "Object o = new Reader(); System.out.println(o instanceof Readable);", "interface Readable {} static class Reader implements Readable {}", "true");
jm!(cast_interface_false, "Object o = new Reader2(); System.out.println(o instanceof Readable2);", "interface Readable2 {} static class Reader2 {}", "false");
jm!(instanceof_chain, "Object o = new Dog(); System.out.println((o instanceof Animal) && ((Animal)o instanceof Pet));", "interface Pet {} static class Animal {} static class Dog extends Animal implements Pet {}", "true");
jm!(cast_chain, "Object o = new D(); D d = (D)((B)o); System.out.println(d.value);", "static class A { int value = 1; } static class B extends A { int value = 2; } static class C extends B {} static class D extends C { int value = 3; }", "3");
jm!(instanceof_before_cast, "Object o = new String(\"x\"); if (o instanceof String) { System.out.println(((String)o).toUpperCase()); }", "", "X");
jm!(instanceof_in_loop, "Object[] arr = {\"x\", 1, new String(\"y\")}; int c = 0; for (Object o : arr) { if (o instanceof String) c++; } System.out.println(c);", "", "2");
jm!(interface_default_instanceof, "System.out.println(Sample.make() instanceof Marker);", "interface Marker {} static class Sample { static Object make() { return new Holder(); } static class Holder implements Marker {} }", "true");
jm!(instanceof_after_cast, "Object o = 3; if (o instanceof Integer) { Integer i = (Integer)o; System.out.println(i + 1); }", "", "4");
jm!(cast_in_ternary, "Object o = 2; System.out.println(o instanceof Integer ? ((Integer)o) : 0);", "", "2");
jm!(cast_from_super, "Object o = new Base(3); Base b = (Base)o; System.out.println(b.value);", "static class Base { int value; Base(int value){ this.value=value; } }", "3");
jm!(cast_reflexive, "System.out.println(new Obj() instanceof Obj);", "static class Obj {}", "true");
jm!(cast_to_same_type, "System.out.println(((Obj2)new Obj2()).value);", "static class Obj2 { int value = 2; }", "2");
jm!(instanceof_final_type, "System.out.println((new X() instanceof X));", "static final class X {}", "true");
jm!(array_instanceof_true, "Object o = new int[]{1,2,3}; System.out.println(o instanceof int[]);", "", "true");
jm!(array_instanceof_false, "Object o = new int[]{1,2,3}; System.out.println(o instanceof String[]);", "", "false");
jm!(cast_to_array, "Object o = new String[]{\"a\"}; String[] a = (String[])o; System.out.println(a.length);", "", "1");
jm!(cast_in_try, "Object x = 1; try { int v = (Integer)x + 1; System.out.println(v); } catch (ClassCastException e) { System.out.println(\"bad\"); }", "", "bad");
jm!(instanceof_parent, "System.out.println(new Sub(\"x\") instanceof Base);", "static class Base {} static class Sub extends Base { Sub(String s) {} }", "true");
jm!(cast_mismatch_with_derived, "Object x = new Dog2(); if (x instanceof Animal2) { System.out.println(((Animal2)x).name()); } else { System.out.println(\"no\"); }", "interface Animal2 { String name(); } static class Dog2 implements Animal2 { public String name() { return \"dog\"; } }", "dog");
jm!(instanceof_multiple, "Object x = new Multi(); System.out.println((x instanceof A2) + "," + (x instanceof B2) + "," + (x instanceof C2));", "interface A2 {} interface B2 {} static class Multi implements A2 {}", "true,false,false");
jm!(cast_with_number, "Object x = Integer.valueOf(7); int y = ((Number)x).intValue(); System.out.println(y);", "", "7");
jm!(instanceof_after_assignment, "Object x = new HolderObj(); HolderObj h = (HolderObj)x; System.out.println(h.value);", "static class HolderObj { int value = 9; }", "9");
jm!(cast_null, "Object x = null; if (x instanceof String) System.out.println(1); else System.out.println(0);", "", "0");
jm!(instanceof_for_switch_like, "Object x = new Integer(1); String out = x instanceof Integer ? \"int\" : \"other\"; System.out.println(out);", "", "int");

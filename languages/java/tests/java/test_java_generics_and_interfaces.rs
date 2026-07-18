use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(generic_box_holds_int, "Box<Integer> b = new Box<>(4); System.out.println(b.value);", "static class Box<T> { T value; Box(T value) { this.value = value; } }", "4");
jm!(generic_method_identity, "System.out.println(Utils.id(5));", "static class Utils { static <T> T id(T v) { return v; } }", "5");
jm!(
    bounded_number_generic,
    "System.out.println(MathUtil.scale((short)2));",
    "static class MathUtil { static int scale(Number n) { return n.intValue() * 2; } }",
    "4"
);
jm!(generic_pair_value, "Pair<String> p = new Pair<>(\"a\", \"b\"); System.out.println(p.first + p.second);", "static class Pair<T> { T first; T second; Pair(T a, T b) { first = a; this.second = b; } }", "ab");
jm!(generic_constructor_and_static_dispatch, "Box<Integer> b = new Box<>(3); System.out.println(b.value);", "static class Box<T> { T value; Box(T value){this.value = value;} }", "3");
jm!(
    interface_generic_contract,
    "List l = new List<Integer>(); l.add(3); System.out.println(l.value());",
    "interface Box<T> { void add(T v); int value(); } static class List<T> implements Box<T> { int count = 0; public void add(T v) { count++; } public int value() { return count; } }",
    "1"
);
jm!(generic_interface_default, "System.out.println(new Node().name());", "interface Named { default String name() { return \"base\"; } } static class Node implements Named {}", "base");
jm!(interface_static_member, "System.out.println(Factory.version());", "interface Factory { static int version() { return 7; } }", "7");
jm!(
    functional_interface_single_method,
    "MathOp op = (int x) -> x + 1; System.out.println(op.apply(3));",
    "@FunctionalInterface interface MathOp { int apply(int x); }",
    "4"
);
jm!(
    generic_interface_inheritance,
    "Item<String> s = new Item<>(\"ok\"); System.out.println(s.get());",
    "interface Provider<T> { T get(); } static class Item<T> implements Provider<T> { T v; Item(T v){this.v=v;} public T get(){ return v; } }",
    "ok"
);
jm!(
    generic_bounded_comparison,
    "Comparator c = new Comparator<Integer>(); System.out.println(c.less(1, 2));",
    "static class Comparator { int less(Number a, Number b) { return a.intValue() < b.intValue() ? 1 : 0; } }",
    "1"
);
jm!(
    interface_multiple_default_methods,
    "System.out.println(new Blend().tag());",
    "interface A { default String tag() { return \"a\"; } } interface B { default String tag() { return \"b\"; } } static class Blend implements A, B { public String tag() { return A.super.tag() + B.super.tag(); } }",
    "ab"
);
jm!(wildcard_extends, "Object[] boxed = {Boxer.box(1), Boxer.box(\"x\")}; System.out.println(boxed.length);", "static class Boxer { static <T> Box<T> box(T v) { return new Box<T>(v); } } static class Box<T> { T v; Box(T v){this.v=v;} }", "2");
jm!(generic_array_list_emulation, "PairList p = new PairList(1,2); System.out.println(p.first() + p.second());", "static class PairList<T> { T a; T b; PairList(T a, T b){this.a=a; this.b=b;} T first(){ return a; } T second(){ return b; } } static class PairList extends PairList<Integer> { PairList(Integer a, Integer b){ super(a,b); } }", "3");
jm!(generic_method_chain, "System.out.println(Util.wrapAndUnwrap(2));", "static class Util { static <T> Holder<T> wrap(T x) { return new Holder<>(x); } static int wrapAndUnwrap(Integer x) { return wrap(x).v; } static class Holder<T> { T v; Holder(T v){this.v=v;} } }", "2");
jm!(array_of_generic_boxes, "Box<Integer>[] boxes = new Box[1]; boxes[0] = new Box<>(9); System.out.println(boxes[0].v);", "static class Box<T> { T v; Box(T v){this.v=v;} }", "9");
jm!(generic_override_in_subclass, "Derived d = new Derived(); System.out.println(d.pick(1));", "static class Base<T> { T pick(T v){ return v; } } static class Derived extends Base<Integer> { Integer pick(Integer v) { return v + 1; } }", "2");
jm!(generic_interface_with_default_impl, "System.out.println(new Store().value());", "interface Getter { default int value() { return 1; } } static class Store implements Getter {}", "1");
jm!(
    generic_bound_compare,
    "System.out.println(Compares.less(new Num(3), new Num(5)));",
    "static class Num implements Comparable<Num> { int v; Num(int v) { this.v = v; } public int compareTo(Num o) { return v - o.v; } } static class Compares { static <T extends Comparable<T>> int less(T a, T b) { return a.compareTo(b) < 0 ? 1 : 0; } }",
    "1"
);
jm!(generic_void_method, "Util.consume(7); System.out.println(1);", "static class Util { static <T> void consume(T value) {} }", "1");
jm!(generic_two_params, "Pair2 p = new Pair2(1, \"x\"); System.out.println(p.first + p.second);", "static class Pair2<A, B> { A first; B second; Pair2(A a, B b){first=a; second=b;} }", "1x");
jm!(interface_as_property, "Worker w = new Worker(); System.out.println(w.run());", "interface Runner { default String run() { return \"base\"; } } static class Worker implements Runner { public String run() { return Runner.super.run() + \"-done\"; } }", "base-done");
jm!(generic_factory_method, "System.out.println(Factory.create(5));", "static class Factory { static <T> T id(T v) { return v; } static int create(int n){ return id(n); } }", "5");
jm!(bounded_generic_function, "System.out.println(Nums.min(9, 4));", "static class Nums { static <T extends Number> int min(T a, T b) { return a.doubleValue() <= b.doubleValue() ? a.intValue() : b.intValue(); } }", "4");
jm!(generic_super_method_chain, "System.out.println(Chain.start(\"a\").append(\"b\").label());", "static class Chain<T> { T value; Chain(T value){this.value=value;} Chain<T> append(T next) { value = (T) (value.toString() + next.toString()); return this; } String label() { return value.toString(); } static Chain<String> start(String v) { return new Chain<>(v); } }", "ab");

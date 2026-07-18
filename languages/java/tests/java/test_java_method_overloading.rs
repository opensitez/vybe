use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(overload_arity, "System.out.println(Mather.add(3) + \",\" + Mather.add(2, 4));", "static class Mather { static int add(int a) { return a; } static int add(int a, int b) { return a + b; } }", "3,6");
jm!(overload_types_int_double, "System.out.println(Chooser.pick(2) + \",\" + Chooser.pick(1.5));", "static class Chooser { static String pick(int v) { return \"i\"; } static String pick(double v) { return \"d\"; } }", "i,d");
jm!(overload_boolean_string, "System.out.println(Dispatcher.tag(true) + \",\" + Dispatcher.tag(\"x\"));", "static class Dispatcher { static String tag(boolean v) { return \"b\"; } static String tag(String v) { return \"s\"; } }", "b,s");
jm!(overload_instance_arity, "System.out.println(Tools.echo(1) + \",\" + Tools.echo(1, 2));", "static class Tools { String echo(int a) { return \"x\"; } String echo(int a, int b) { return \"xx\"; } }", "x,xx");
jm!(overload_return_not_used, "System.out.println(Chain.value(1,2));", "static class Chain { static int value(int a) { return 1; } static int value(int a, int b, int c) { return 3; } }", "1");
jm!(overload_boolean_boxed, "System.out.println(Mode.level(true) + \",\" + Mode.level(Boolean.valueOf(false)));", "static class Mode { static int level(boolean v) { return 1; } static int level(Boolean v) { return 2; } }", "1,2");
jm!(constructor_overload, "System.out.println(new Bag().n + \",\" + new Bag(2).n);", "static class Bag { int n; Bag() { n = 1; } Bag(int n) { this.n = n; } }", "1,2");
jm!(method_chain_overloaded, "System.out.println(new Pair().sum(1) + \",\" + new Pair().sum(1,2));", "static class Pair { int sum(int a) { return a; } int sum(int a, int b) { return a + b; } }", "1,3");
jm!(overload_uses_array, "System.out.println(Arrayish.of(1) + \",\" + Arrayish.of(new int[]{1,2}));", "static class Arrayish { static int of(int a) { return 1; } static int of(int[] values) { return values.length; } }", "1,2");
jm!(overload_stringable, "System.out.println(View.label(\"a\") + \",\" + View.label(new Object()));", "static class View { static String label(String s) { return \"str\"; } static String label(Object o) { return \"obj\"; } }", "str,obj");
jm!(overload_conflict_default, "System.out.println(Conf.toInt(1) + \",\" + Conf.toInt(\"2\"));", "static class Conf { static int toInt(int v) { return v; } static int toInt(String s) { return Integer.parseInt(s); } }", "1,2");
jm!(instance_overload_mixed, "System.out.println(new Box().id(1) + \",\" + new Box().id(\"x\"));", "static class Box { String id(int v) { return \"i\"+v; } String id(String v) { return \"s\"+v; } }", "i1,sx");
jm!(overload_with_parentheses, "System.out.println(MathUtil.make(1) + \",\" + MathUtil.make(1,2));", "static class MathUtil { static int make(int x) { return x; } static int make(int x, int y) { return x + y; } }", "1,3");
jm!(overload_with_char_array, "System.out.println(Code.of('a') + \",\" + Code.of('a','b'));", "static class Code { static int of(char c) { return 1; } static int of(char c1, char c2) { return 2; } }", "1,2");
jm!(overload_same_type_variants, "System.out.println(Combo.choose(1, true) + \",\" + Combo.choose(1, 2));", "static class Combo { static int choose(int a, boolean b) { return 1; } static int choose(int a, int b) { return 2; } }", "1,2");
jm!(overload_null_dispatch, "System.out.println(Flow.kind(null) + \",\" + Flow.kind(\"x\"));", "static class Flow { static String kind(Object o) { return \"obj\"; } static String kind(String o) { return \"str\"; } }", "obj,str");
jm!(overload_double_call, "System.out.println(Form.value(1) + \",\" + Form.value(1.0) + \",\" + Form.value(1f));", "static class Form { static String value(double v) { return \"d\"; } static String value(float v) { return \"f\"; } }", "d,f,f");
jm!(overload_through_subclass, "System.out.println((new Child()).label(1) + \",\" + (new Child()).label(1,2));", "static class Parent { int label(int v) { return 1; } int label(int a, int b) { return 2; } } static class Child extends Parent {}", "1,2");
jm!(overload_instance_and_static, "System.out.println(Over.staticOp(1) + \",\" + new Over().staticOp(1));", "static class Over { static int staticOp(int a) { return 1; } int staticOp(int a, int b) { return 2; } }", "1,2");
jm!(overload_constructor_chain, "System.out.println(new Maker().x + \",\" + new Maker(2).x);", "static class Maker { int x; Maker() { this(1); } Maker(int x) { this.x = x; } }", "1,2");
jm!(overload_three_methods, "System.out.println(Util.merge(1) + \",\" + Util.merge(1,2) + \",\" + Util.merge(1,2,3));", "static class Util { static int merge(int a) { return a; } static int merge(int a, int b) { return a+b; } static int merge(int a, int b, int c) { return a+b+c; } }", "1,3,6");
jm!(overload_with_boolean_object, "System.out.println(Flag.pick(1) + \",\" + Flag.pick(Boolean.TRUE));", "static class Flag { static int pick(int v) { return 1; } static int pick(Boolean v) { return 2; } }", "1,2");
jm!(overload_same_arity_primitive, "System.out.println(Mode2.select(1) + \",\" + Mode2.select(1.0));", "static class Mode2 { static String select(int v) { return \"i\"; } static String select(double v) { return \"d\"; } }", "i,d");
jm!(overload_local_default, "System.out.println(Resolver.eval(1) + \",\" + Resolver.eval(\"1\"));", "static class Resolver { static int eval(int v) { return v; } static int eval(String v) { return Integer.parseInt(v); } }", "1,1");
jm!(overload_in_interface_context, "System.out.println(Worker.pick(1) + \",\" + Worker.pick(\"1\"));", "interface Work { static int pick(int n) { return 0; } } static class Worker implements Work { static int pick(int n) { return 3; } static int pick(String s) { return Integer.parseInt(s); } }", "3,1");

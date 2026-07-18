use crate::helpers::run_in_main;

macro_rules! jm {
    ($name:ident, $src:expr, $types:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_in_main($src, $types), vec![$expected]);
        }
    };
}

jm!(sum_varargs_empty, "System.out.println(Var.total());", "static class Var { static int total(int... values) { int s = 0; for (int v : values) { s += v; } return s; } }", "0");
jm!(sum_varargs_three, "System.out.println(Var.total(1, 2, 3));", "static class Var { static int total(int... values) { int s = 0; for (int v : values) { s += v; } return s; } }", "6");
jm!(sum_varargs_plus, "System.out.println(Var.add(2, Var.total(3,4)));", "static class Var { static int total(int... values) { int s = 0; for (int v : values) { s += v; } return s; } static int add(int base, int... values) { return base + total(values); } }", "9");
jm!(single_argument_no_var, "System.out.println(Var.total(4));", "static class Var { static int total(int... values) { int s = 0; for (int v : values) { s += v; } return s; } }", "4");
jm!(overloaded_fixed_and_varargs, "System.out.println(Var.echo(1) + "," + Var.echo(1, 2));", "static class Var { static int echo(int a) { return 1; } static int echo(int a, int... rest) { return 1 + rest.length; } }", "1,2");
jm!(string_varargs_join, "System.out.println(Joiner.join("x", "y", "z"));", "static class Joiner { static String join(String first, String... rest) { String out = first; for(String p : rest){ out += p; } return out; } }", "xyz");
jm!(varargs_instance, "System.out.println(Counter.from(1,2,3).sum());", "static class Counter { static Counter from(int... values) { Counter c = new Counter(); for (int v : values) { c.total += v; } return c; } int sum() { return total; } int total = 0; }", "6");
jm!(varargs_instance_add, "System.out.println(new Counter2().with(4).total);", "static class Counter2 { int total = 0; Counter2 with(int... values) { for (int v : values) { total += v; } return this; } }", "4");
jm!(empty_varargs_array, "System.out.println(Aggregator.of(new int[]{}));", "static class Aggregator { static int of(int... values) { int s = 0; for (int v : values) { s += v; } return s; } }", "0");
jm!(varargs_with_cast, "System.out.println(Var2.concat("a", 1, 2));", "static class Var2 { static String concat(String p, int... values) { String s = p; for (int v: values) { s += v; } return s; } }", "a12");
jm!(boolean_varargs, "System.out.println(Checker.flags(true, false, true));", "static class Checker { static int flags(boolean... values) { int s = 0; for (boolean v : values) { s += v ? 1 : 0; } return s; } }", "2");
jm!(double_varargs, "System.out.println(MathPack.sum(1.5, 2.5, 3.0));", "static class MathPack { static int sum(double... values) { int s = 0; for (double v : values) { s += (int)v; } return s; } }", "6");
jm!(varargs_and_overload1, "System.out.println(Mixer.pick(1) + "," + Mixer.pick(1, 2));", "static class Mixer { static String pick(int a) { return "one"; } static String pick(int a, int... rest) { return "many"; } }", "one,many");
jm!(varargs_and_overload2, "System.out.println(Mixer2.pick(1) + "," + Mixer2.pick(1, 2));", "static class Mixer2 { static String pick(int a) { return "one"; } static String pick(int a, int b) { return "pair"; } static String pick(int a, int b, int c) { return "trip"; } }", "one,pair");
jm!(char_array_varargs, "System.out.println(CharVar.text('a', 'b', 'c'));", "static class CharVar { static String text(char... values) { String s = ""; for (char c : values) { s += c; } return s; } }", "abc");
jm!(varargs_as_array_reference, "System.out.println(ArrVar.total(1,2,3));", "static class ArrVar { static int total(int... values) { int s=0; for (int v : values) { s += v; } return s; } }", "6");
jm!(constructor_varargs, "System.out.println(new Packet().count);", "static class Packet { int count; Packet(int... values) { for (int v : values) { count += v; } } Packet(){ this(1,2,3); } }", "6");
jm!(varargs_in_instance_method, "System.out.println(new Series().sum(1, 2, 3));", "static class Series { int sum(int... values) { int s=0; for (int v : values) { s += v; } return s; } }", "6");
jm!(generic_vararg_like_object, "System.out.println(ObjectVar.concat("A", 1, 2));", "static class ObjectVar { static String concat(Object first, Object... rest) { String s = String.valueOf(first); for (Object o : rest) { s += String.valueOf(o); } return s; } }", "A12");
jm!(varargs_multiple_calls, "System.out.println(Seq.total(1) + "," + Seq.total(1,2) + "," + Seq.total(1,2,3));", "static class Seq { static int total(int first, int... rest) { int s = first; for (int v : rest) { s += v; } return s; } }", "1,3,6");
jm!(varargs_large, "System.out.println(Range.add(1,2,3,4,5));", "static class Range { static int add(int... values) { int s=0; for (int v : values) { s += v; } return s; } }", "15");
jm!(varargs_with_constant, "System.out.println(Defaults.join(3, 4, 5));", "static class Defaults { static int join(int base, int... rest) { return total(base, rest); } static int total(int base, int... values) { int s = base; for (int v : values) { s += v; } return s; } }", "12");
jm!(nested_varargs_calls, "System.out.println(Caller.call(7));", "static class Caller { static int call(int v) { return inner(v, 1, 2, 3); } static int inner(int base, int... rest) { int s = base; for (int i : rest) { s += i; } return s; } }", "13");
jm!(varargs_and_non_varargs_conflict, "System.out.println(Pivot.a(1) + "," + Pivot.a(1,2));", "static class Pivot { static int a(int a) { return 1; } static int a(int a, int b) { return 2; } static int a(int a, int b, int c) { return 3; } }", "1,2");
jm!(call_with_explicit_array, "System.out.println(Maker.total(new int[]{2,3,4}));", "static class Maker { static int total(int... values) { int s=0; for(int v : values) { s += v; } return s; } }", "9");
jm!(varargs_method_reference, "System.out.println(Emitter.emit(1,2,3).length());", "static class Emitter { static String emit(int... values) { return String.valueOf(values.length); } }", "3");

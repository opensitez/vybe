// vybe-test: csharp/csharp_nested_partial_types/nested_class_inside_generic_outer_type_uses_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

var wrapper = new Box<int>.Wrapper { Value = 9 }
;
__P((wrapper.Value).ToString());
__Check("9");

class Box<T> {
    public class Wrapper {
        public T Value { get; set; }
    }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}

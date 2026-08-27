// vybe-test: csharp/csharp_generic_inference_calls/generic_nested_type_shares_outer_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

var built = new Outer<string>().Build("nested");
__P((built.Value).ToString());
__Check("nested");

class Outer<T> {
    public class Inner {
        public T Value;
    }
    public Inner Build(T value) {
        return new Inner { Value = value };
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

// vybe-test: csharp/csharp_generic_inference_calls/generic_class_method_infers_from_arguments
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

var numbers = new Box<int>(5);
__P((numbers.Get()).ToString());
__Check("5");

class Box<T> {
    public T Value;
    public Box(T value) { Value = value; }
    public T Get() { return Value; }
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

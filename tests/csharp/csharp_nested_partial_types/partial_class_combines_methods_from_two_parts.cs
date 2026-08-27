// vybe-test: csharp/csharp_nested_partial_types/partial_class_combines_methods_from_two_parts
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

var worker = new Worker();
__P((worker.First()).ToString());
__P((worker.Second()).ToString());
__Check("one\ntwo");

partial class Worker {
    public string First() { return "one"; }
}

partial class Worker {
    public string Second() { return "two"; }
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

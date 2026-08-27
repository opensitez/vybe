// vybe-test: csharp/csharp_default_interface_methods/default_interface_method_visible_through_interface_typed_reference
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

using static __Harness;

ICounter counter = new Counter { Value = 4 }
;
__P((counter.Next()).ToString());
__Check("5");

interface ICounter {
    int Value { get; }
    int Next() { return Value + 1; }
}

class Counter : ICounter {
    public int Value { get; set; }
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

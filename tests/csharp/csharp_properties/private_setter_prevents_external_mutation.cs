// vybe-test: csharp/csharp_properties/private_setter_prevents_external_mutation
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

using static __Harness;

var c = new Counter();
c.Increment();
c.Increment();
__P((c.Count).ToString());
__Check("2");

class Counter {
    public int Count { get; private set; }
    public void Increment() => Count++;
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

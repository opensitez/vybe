// vybe-test: csharp/csharp_properties_accessors/private_setter_property_changes_through_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var counter = new Counter();
counter.Increment();
counter.Increment();
__P((counter.Value).ToString());
__Check("2");

class Counter {
    public int Value { get; private set; }
    public void Increment() { Value++; }
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

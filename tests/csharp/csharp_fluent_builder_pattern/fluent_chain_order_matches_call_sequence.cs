// vybe-test: csharp/csharp_fluent_builder_pattern/fluent_chain_order_matches_call_sequence
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

using static __Harness;

var trace = new Trace().Step("a").Step("b").Step("c");
__P((trace.Read()).ToString());
__Check("abc");

class Trace {
    string log = "";
    public Trace Step(string name) { log += name; return this; }
    public string Read() { return log; }
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

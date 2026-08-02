// vybe-test: csharp/csharp_fluent_builder_pattern/fluent_chain_order_matches_call_sequence
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    string log = "";
    public Trace Step(string name) { log += name; return this; }
    public string Read() { return log; }
}
var trace = new Trace().Step("a").Step("b").Step("c");
__Check((trace.Read()).ToString(), "abc");

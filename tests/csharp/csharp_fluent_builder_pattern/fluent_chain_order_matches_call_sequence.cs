// vybe-test: csharp/csharp_fluent_builder_pattern/fluent_chain_order_matches_call_sequence
// origin: languages/csharp/tests/csharp/test_csharp_fluent_builder_pattern.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Trace {
    string log = "";
    public Trace Step(string name) { log += name; return this; }
    public string Read() { return log; }
}
var trace = new Trace().Step("a").Step("b").Step("c");
__P((trace.Read()).ToString());
__Check("abc");

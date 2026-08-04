// vybe-test: csharp/csharp_virtual_dispatch_semantics/this_constructor_chain_reuses_sibling_constructor_logic
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

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

class Pair {
    public int First;
    public int Second;
    public Pair(int value) : this(value, value) { }
    public Pair(int first, int second) { First = first; Second = second; }
}
var pair = new Pair(9);
__P((pair.First).ToString());
__P((pair.Second).ToString());
__Check("9\n9");

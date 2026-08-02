// vybe-test: csharp/csharp_virtual_dispatch_semantics/this_constructor_chain_reuses_sibling_constructor_logic
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((pair.First).ToString(), "9");
__Check((pair.Second).ToString(), "9");

// vybe-test: csharp/collections_advanced/collection_initializer_dict
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

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

var ages = new Dictionary<string, int> {
    { "Alice", 30 },
    { "Bob", 25 }
};
__P((ages["Alice"]).ToString());
__P((ages.Count).ToString());
__Check("30\n2");

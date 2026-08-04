// vybe-test: csharp/csharp_class_indexers/string_keyed_indexer_stores_and_retrieves_values
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

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

class Bag {
    System.Collections.Generic.Dictionary<string, int> map = new();
    public int this[string key] {
        get { return map[key]; }
        set { map[key] = value; }
    }
}
var bag = new Bag();
bag["count"] = 7;
__P((bag["count"]).ToString());
__Check("7");

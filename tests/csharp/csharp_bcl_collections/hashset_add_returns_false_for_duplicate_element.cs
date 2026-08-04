// vybe-test: csharp/csharp_bcl_collections/hashset_add_returns_false_for_duplicate_element
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

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

var set = new System.Collections.Generic.HashSet<int>();
__P((set.Add(1)).ToString());
__P((set.Add(1)).ToString());
__Check("True\nFalse");

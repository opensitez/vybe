// vybe-test: csharp/collections_advanced/hashset_basic
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

var set = new HashSet<int> { 1, 2, 3, 2, 1 };
__P((set.Count).ToString());
__P((set.Contains(2)).ToString());
__P((set.Contains(5)).ToString());
__Check("3\nTrue\nFalse");

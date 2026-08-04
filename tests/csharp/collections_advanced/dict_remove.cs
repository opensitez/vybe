// vybe-test: csharp/collections_advanced/dict_remove
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

var dict = new Dictionary<string, int> { { "a", 1 }, { "b", 2 }, { "c", 3 } };
dict.Remove("b");
__P((dict.Count).ToString());
__P((dict.ContainsKey("b")).ToString());
__Check("2\nFalse");

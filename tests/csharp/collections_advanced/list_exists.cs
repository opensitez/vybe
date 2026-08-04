// vybe-test: csharp/collections_advanced/list_exists
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

var list = new List<string> { "apple", "banana", "cherry" };
__P((list.Exists(s => s == "banana")).ToString());
__P((list.Exists(s => s == "grape")).ToString());
__Check("True\nFalse");

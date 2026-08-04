// vybe-test: csharp/collections_advanced/hashset_add_remove
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

var set = new HashSet<string>();
set.Add("apple");
set.Add("banana");
set.Add("apple");
__P((set.Count).ToString());
set.Remove("apple");
__P((set.Count).ToString());
__P((set.Contains("apple")).ToString());
__Check("2\n1\nFalse");

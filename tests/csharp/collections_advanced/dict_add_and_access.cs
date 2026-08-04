// vybe-test: csharp/collections_advanced/dict_add_and_access
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

var dict = new Dictionary<string, int>();
dict.Add("one", 1);
dict.Add("two", 2);
dict.Add("three", 3);
__P((dict["two"]).ToString());
__P((dict.Count).ToString());
__Check("2\n3");

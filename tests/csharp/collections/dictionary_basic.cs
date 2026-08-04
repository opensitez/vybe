// vybe-test: csharp/collections/dictionary_basic
// origin: languages/csharp/tests/csharp/test_collections.rs

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
        dict.Add("x", 10);
        dict.Add("y", 20);
        __P((dict["x"]).ToString());
        __P((dict["y"]).ToString());
__Check("10\n20");

// vybe-test: csharp/csharp_icomparable_sorting/comparer_default_sorts_strings_lexicographically
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

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

var list = new System.Collections.Generic.List<string>{"banana","apple","cherry"};
list.Sort(System.StringComparer.Ordinal);
__P((list[0]).ToString());
__Check("apple");

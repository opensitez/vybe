// vybe-test: csharp/csharp_ranges_indices/range_on_string_returns_substring
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

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

string s="hello world"; __P((s[6..]).ToString());
__Check("world");

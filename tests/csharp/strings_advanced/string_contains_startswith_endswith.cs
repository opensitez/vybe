// vybe-test: csharp/strings_advanced/string_contains_startswith_endswith
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

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

string s = "Hello World";
__P((s.Contains("lo Wo")).ToString());
__P((s.StartsWith("Hello")).ToString());
__P((s.EndsWith("World")).ToString());
__P((s.StartsWith("World")).ToString());
__Check("True\nTrue\nTrue\nFalse");

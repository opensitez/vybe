// vybe-test: csharp/csharp_parsing_formatting/string_format_replaces_indexed_placeholders
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

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

__P((string.Format("{0}-{1}", "A", 3)).ToString());
__Check("A-3");

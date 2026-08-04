// vybe-test: csharp/csharp_strings_ext/string_intern_returns_reference_equal_for_same_literal_sequence
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

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

string a = string.Intern("shared");
string b = string.Intern("shared");
__P((object.ReferenceEquals(a, b)).ToString());
__Check("True");

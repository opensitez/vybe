// vybe-test: csharp/csharp_string_culture/invariant_culture_tostring_for_double_uses_dot_separator
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

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

double d=1.5;
__P((d.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToString());
__Check("1.5");

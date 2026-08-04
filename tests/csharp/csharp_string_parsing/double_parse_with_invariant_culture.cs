// vybe-test: csharp/csharp_string_parsing/double_parse_with_invariant_culture
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

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

var d=double.Parse("3.14",System.Globalization.CultureInfo.InvariantCulture);
__P((d).ToString());
__Check("3.14");

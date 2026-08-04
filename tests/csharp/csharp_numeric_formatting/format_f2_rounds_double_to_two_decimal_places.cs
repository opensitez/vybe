// vybe-test: csharp/csharp_numeric_formatting/format_f2_rounds_double_to_two_decimal_places
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

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

__P(((3.14159).ToString("F2")).ToString());
__Check("3.14");

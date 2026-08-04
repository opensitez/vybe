// vybe-test: csharp/csharp_type_casting/double_to_int_truncates_fractional_part
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

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

double d = 9.9; int n = (int)d; __P((n).ToString());
__Check("9");

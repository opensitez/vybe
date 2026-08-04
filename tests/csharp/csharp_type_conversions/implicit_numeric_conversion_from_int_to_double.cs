// vybe-test: csharp/csharp_type_conversions/implicit_numeric_conversion_from_int_to_double
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

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

int count = 7; double total = count; __P((total).ToString());
__Check("7");

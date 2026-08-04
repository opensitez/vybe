// vybe-test: csharp/csharp_type_conversions/convert_class_creates_double_from_integer
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

__P((System.Convert.ToDouble(5)).ToString());
__Check("5");

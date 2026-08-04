// vybe-test: csharp/csharp_conversion_methods/convert_to_double_from_int
// origin: languages/csharp/tests/csharp/test_csharp_conversion_methods.rs

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

double d=System.Convert.ToDouble(7);
__P((d).ToString());
__Check("7");

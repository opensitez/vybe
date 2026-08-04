// vybe-test: csharp/csharp_numeric_precision/float_is_32_bit_and_less_precise_than_double
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

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

float f=1.0f/3.0f;
double d=1.0/3.0;
__P((f==(float)d).ToString());
__Check("True");

// vybe-test: csharp/csharp_integer_arithmetic/reassignment_after_arithmetic_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

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

int value = 1; value = value + 2 + 3; __P((value).ToString());
__Check("6");

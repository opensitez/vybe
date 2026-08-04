// vybe-test: csharp/csharp_integer_arithmetic/arithmetic_with_multiple_variables_in_one_expression
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

int a = 2; int b = 3; int c = 4; __P((a + b * c).ToString());
__Check("14");

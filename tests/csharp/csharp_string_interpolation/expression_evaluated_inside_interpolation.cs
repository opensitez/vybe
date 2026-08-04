// vybe-test: csharp/csharp_string_interpolation/expression_evaluated_inside_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

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

int a=3,b=4; __P(($"{a}+{b}={a+b}").ToString());
__Check("3+4=7");

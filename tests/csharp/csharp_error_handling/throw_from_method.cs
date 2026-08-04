// vybe-test: csharp/csharp_error_handling/throw_from_method
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

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

int Divide(int a, int b) {
    if (b == 0) throw new Exception("Division by zero");
    return a / b;
}
try {
    __P((Divide(10, 2)).ToString());
    __P((Divide(10, 0)).ToString());
} catch (Exception e) {
    __P((e.Message).ToString());
}
__Check("5\nDivision by zero");

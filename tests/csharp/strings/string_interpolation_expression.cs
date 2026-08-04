// vybe-test: csharp/strings/string_interpolation_expression
// origin: languages/csharp/tests/csharp/test_strings.rs

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

var x = 3;
        var y = 4;
        __P(($"sum is {x + y}").ToString());
__Check("sum is 7");

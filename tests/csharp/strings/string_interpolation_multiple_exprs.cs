// vybe-test: csharp/strings/string_interpolation_multiple_exprs
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

var a = "Alice";
        var age = 30;
        __P(($"{a} is {age} years old").ToString());
__Check("Alice is 30 years old");

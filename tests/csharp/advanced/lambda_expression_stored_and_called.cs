// vybe-test: csharp/advanced/lambda_expression_stored_and_called
// origin: languages/csharp/tests/csharp/test_advanced.rs

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

var twice = x => x * 2;
        var result = twice(5);
        __P((result).ToString());
__Check("10");

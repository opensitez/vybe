// vybe-test: csharp/type_features/lambda_expression_foreach
// origin: languages/csharp/tests/csharp/test_type_features.rs

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

var arr = new int[] { 1, 2, 3, 4, 5 };
        var sum = 0;
        foreach (var x in arr) { sum = sum + x; }
        __P((sum).ToString());
__Check("15");

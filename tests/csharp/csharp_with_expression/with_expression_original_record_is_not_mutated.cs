// vybe-test: csharp/csharp_with_expression/with_expression_original_record_is_not_mutated
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

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

record Point(int X, int Y);
var origin = new Point(1, 2);
var moved = origin with { X = 10 };
__P((origin.X).ToString());
__Check("1");

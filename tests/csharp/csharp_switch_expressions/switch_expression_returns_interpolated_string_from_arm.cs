// vybe-test: csharp/csharp_switch_expressions/switch_expression_returns_interpolated_string_from_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

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

var score = 87; __P((score switch { >= 90 => $"A:{score}", >= 80 => $"B:{score}", _ => $"C:{score}" }).ToString());
__Check("B:87");

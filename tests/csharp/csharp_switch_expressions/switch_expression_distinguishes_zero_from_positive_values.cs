// vybe-test: csharp/csharp_switch_expressions/switch_expression_distinguishes_zero_from_positive_values
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

var x = 0; __P((x switch { < 0 => "neg", 0 => "zero", > 0 => "pos" }).ToString());
__Check("zero");

// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_small_integer_constant
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

var x = 2; __P((x switch { 1 => "one", 2 => "two", _ => "other" }).ToString());
__Check("two");

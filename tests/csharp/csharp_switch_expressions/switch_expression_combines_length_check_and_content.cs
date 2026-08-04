// vybe-test: csharp/csharp_switch_expressions/switch_expression_combines_length_check_and_content
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

var text = "tool"; __P((text switch { string s when s.Length == 4 => "len4", string s => s, _ => "none" }).ToString());
__Check("len4");

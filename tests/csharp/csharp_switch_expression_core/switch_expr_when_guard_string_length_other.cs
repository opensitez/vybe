// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_guard_string_length_other
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

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

var s="hi"; __P((s switch{string t when t.Length==4=>"len4",string t=>t.Length.ToString(),_=>"0"}).ToString());
__Check("2");

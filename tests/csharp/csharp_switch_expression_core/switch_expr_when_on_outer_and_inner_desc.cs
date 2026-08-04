// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_on_outer_and_inner_desc
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

var pair=(5,2); __P((pair switch{(var a,var b) when a<b=>"asc",(var a,var b)=>"desc",_=>"?"}).ToString());
__Check("desc");

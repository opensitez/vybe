// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_switch_as_arm_value
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

var tier=2; __P((tier switch{1=>"a",2=>(3 switch{3=>"inner",_=>"outer"}),_=>"?"}).ToString());
__Check("outer");

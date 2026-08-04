// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_arm_switch_default_inner
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

string Outer(int n)=>n switch{1=>5 switch{5=>"five",_=>"not-five"},_=>"rest"}; __P((Outer(1)).ToString());
__Check("not-five");

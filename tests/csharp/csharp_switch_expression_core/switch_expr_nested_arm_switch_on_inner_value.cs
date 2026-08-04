// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_arm_switch_on_inner_value
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

string Outer(int n)=>n switch{1=>1 switch{1=>"one-one",_=>"one-other"},2=>"two",_=>"rest"}; __P((Outer(1)).ToString());
__Check("one-one");

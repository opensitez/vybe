// vybe-test: csharp/csharp_switch_expression_core/switch_expr_double_nested_in_arm_result
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

int Pick(int a,int b)=>a switch{1=>b switch{2=>10,3=>20,_=>0},_=>-1}; __P((Pick(1,3)).ToString());
__Check("20");

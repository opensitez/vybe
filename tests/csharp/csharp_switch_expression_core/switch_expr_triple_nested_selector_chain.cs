// vybe-test: csharp/csharp_switch_expression_core/switch_expr_triple_nested_selector_chain
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

int Depth(int n)=>n switch{1=>2 switch{2=>3 switch{3=>9,_=>0},_=>0},_=>0}; __P((Depth(1)).ToString());
__Check("0");

// vybe-test: csharp/csharp_conditional_expressions/ternary_nested_three_way_comparison
// origin: languages/csharp/tests/csharp/test_csharp_conditional_expressions.rs

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

int n=0;
__P((n>0?"pos":n<0?"neg":"zero").ToString());
__Check("zero");

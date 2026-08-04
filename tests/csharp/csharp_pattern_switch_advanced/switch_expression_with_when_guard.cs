// vybe-test: csharp/csharp_pattern_switch_advanced/switch_expression_with_when_guard
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

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

string Classify(int n)=>n switch{
    int x when x<0=>"negative",
    0=>"zero",
    int x when x%2==0=>"even",
    _=>"odd"};
__P((Classify(-5)).ToString());
__P((Classify(0)).ToString());
__P((Classify(4)).ToString());
__P((Classify(7)).ToString());
__Check("negative\nzero\neven\nodd");

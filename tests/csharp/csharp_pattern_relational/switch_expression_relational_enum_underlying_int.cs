// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_enum_underlying_int
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

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

enum Tier { Low=1, Mid=5, High=10 } var t=Tier.Mid; __P((t switch{>=Tier.Mid=>"up",_=>"down"}).ToString());
__Check("up");

// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_shift_left_style_combine
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

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

using System; [Flags] enum P{A=1,B=2,C=4} var v=(P)0; v|=P.A; v|=P.C; __P(((int)v).ToString());
__Check("5");

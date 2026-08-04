// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_enum_in_switch_print
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

using System; [Flags] enum P{Read=1,Write=2} P v=P.Read; string s=v.HasFlag(P.Write)?"w":"r"; __P((s).ToString());
__Check("r");

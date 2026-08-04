// vybe-test: csharp/csharp_switch_type_patterns/switch_on_enum_uses_symbolic_case_labels
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

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

enum Tier { Free, Pro }
string Name(Tier tier) => tier switch {
    Tier.Free => "free",
    Tier.Pro => "pro",
    _ => "unknown"
};
__P((Name(Tier.Pro)).ToString());
__Check("pro");

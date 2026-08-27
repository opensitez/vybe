// vybe-test: csharp/csharp_switch_type_patterns/switch_on_enum_uses_symbolic_case_labels
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

string Name(Tier tier) => tier switch {
    Tier.Free => "free",
    Tier.Pro => "pro",
    _ => "unknown"
}
;
__P((Name(Tier.Pro)).ToString());
__Check("pro");

enum Tier { Free, Pro }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}

// vybe-test: csharp/csharp_switch_type_patterns/switch_on_enum_uses_symbolic_case_labels
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Tier { Free, Pro }
string Name(Tier tier) => tier switch {
    Tier.Free => "free",
    Tier.Pro => "pro",
    _ => "unknown"
};
__Check((Name(Tier.Pro)).ToString(), "pro");

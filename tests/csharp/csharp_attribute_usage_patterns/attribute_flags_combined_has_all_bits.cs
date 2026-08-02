// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_combined_has_all_bits
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Flags] enum P{A=1,B=2,C=4} var v=P.A|P.B|P.C; __Check((v.HasFlag(P.A)&&v.HasFlag(P.B)&&v.HasFlag(P.C)).ToString(), "True");

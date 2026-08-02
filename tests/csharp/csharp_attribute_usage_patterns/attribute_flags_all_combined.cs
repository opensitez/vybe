// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_all_combined
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Flags] enum P{A=1,B=2,C=4} __Check(((int)(P.A|P.B|P.C)).ToString(), "7");

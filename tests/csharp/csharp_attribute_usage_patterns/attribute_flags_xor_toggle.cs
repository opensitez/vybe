// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_xor_toggle
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Flags] enum P{A=1,B=2} var v=P.A|P.B; __Check(((int)(v^P.A)).ToString(), "2");

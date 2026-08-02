// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_none_is_zero
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Flags] enum P{None=0,Read=1} __Check(((int)P.None).ToString(), "0");

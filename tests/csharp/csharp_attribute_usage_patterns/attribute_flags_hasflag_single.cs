// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_hasflag_single
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Flags] enum P{Read=1,Write=2} var v=P.Read|P.Write; __Check((v.HasFlag(P.Read)).ToString(), "True");

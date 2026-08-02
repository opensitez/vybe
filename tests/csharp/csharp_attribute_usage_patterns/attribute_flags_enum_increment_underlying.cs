// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_enum_increment_underlying
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Flags] enum P{Read=1,Write=2,Exec=4} __Check((P.Exec>P.Read).ToString(), "True");

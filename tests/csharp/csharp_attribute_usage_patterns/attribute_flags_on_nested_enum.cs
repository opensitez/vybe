// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_flags_on_nested_enum
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Host{[Flags] public enum M{A=1,B=2}} __Check(((int)(Host.M.A|Host.M.B)).ToString(), "3");

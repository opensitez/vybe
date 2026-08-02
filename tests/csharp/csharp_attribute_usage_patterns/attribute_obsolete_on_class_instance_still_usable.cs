// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_on_class_instance_still_usable
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Obsolete("old")] class S{public int N=1;} __Check((new S().N).ToString(), "1");

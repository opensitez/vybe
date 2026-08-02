// vybe-test: csharp/csharp_casting_patterns/as_returns_typed_reference_for_compatible_type
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o="world";
string s=o as string;
__Check((s!=null).ToString(), "True"); __Check((s).ToString(), "world");

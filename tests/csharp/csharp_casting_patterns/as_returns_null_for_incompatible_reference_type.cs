// vybe-test: csharp/csharp_casting_patterns/as_returns_null_for_incompatible_reference_type
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=42;
string s=o as string;
__Check((s==null).ToString(), "True");

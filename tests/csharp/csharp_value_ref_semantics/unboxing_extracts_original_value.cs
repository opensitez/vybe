// vybe-test: csharp/csharp_value_ref_semantics/unboxing_extracts_original_value
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=42; int n=(int)o;
__Check((n).ToString(), "42");

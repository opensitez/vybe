// vybe-test: csharp/csharp_casting_patterns/is_null_constant_pattern_detects_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=null;
__Check((o is null).ToString(), "True");

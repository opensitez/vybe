// vybe-test: csharp/csharp_null_handling/is_null_pattern_matches_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o = null;
__Check((o is null).ToString(), "True");

// vybe-test: csharp/csharp_null_handling/is_not_null_pattern_matches_non_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o = "hello";
__Check((o is not null).ToString(), "True");

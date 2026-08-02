// vybe-test: csharp/csharp_pattern_matching_advanced/constant_pattern_matches_true_boolean
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object value = true; if (value is true) __Check(("yes").ToString(), "yes");

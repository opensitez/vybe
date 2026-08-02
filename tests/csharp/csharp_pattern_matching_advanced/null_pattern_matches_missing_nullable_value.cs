// vybe-test: csharp/csharp_pattern_matching_advanced/null_pattern_matches_missing_nullable_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? value = null; if (value is null) __Check(("missing").ToString(), "missing");

// vybe-test: csharp/csharp_pattern_matching_advanced/relational_pattern_matches_value_in_if_statement
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var score = 91; if (score is >= 90) __Check(("A").ToString(), "A");

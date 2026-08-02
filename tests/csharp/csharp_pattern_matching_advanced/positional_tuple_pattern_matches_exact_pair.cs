// vybe-test: csharp/csharp_pattern_matching_advanced/positional_tuple_pattern_matches_exact_pair
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pair = (2, 3); if (pair is (2, 3)) __Check(("match").ToString(), "match");

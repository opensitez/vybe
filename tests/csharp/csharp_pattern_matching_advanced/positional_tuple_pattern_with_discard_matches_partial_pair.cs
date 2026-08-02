// vybe-test: csharp/csharp_pattern_matching_advanced/positional_tuple_pattern_with_discard_matches_partial_pair
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pair = (2, 9); if (pair is (2, _)) __Check(("left-two").ToString(), "left-two");

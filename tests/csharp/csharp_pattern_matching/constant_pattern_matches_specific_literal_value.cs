// vybe-test: csharp/csharp_pattern_matching/constant_pattern_matches_specific_literal_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 3;
string result = x switch { 1 => "one", 2 => "two", 3 => "three", _ => "other" };
__Check((result).ToString(), "three");

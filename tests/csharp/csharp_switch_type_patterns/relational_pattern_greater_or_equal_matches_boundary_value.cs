// vybe-test: csharp/csharp_switch_type_patterns/relational_pattern_greater_or_equal_matches_boundary_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int score = 100;
string grade = score switch {
    >= 90 => "A",
    >= 80 => "B",
    _ => "C"
};
__Check((grade).ToString(), "A");

// vybe-test: csharp/csharp_pattern_matching/relational_pattern_compares_numeric_range
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int score = 85;
string grade = score switch { >= 90 => "A", >= 80 => "B", >= 70 => "C", _ => "F" };
__Check((grade).ToString(), "B");

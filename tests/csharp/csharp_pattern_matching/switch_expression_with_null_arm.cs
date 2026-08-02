// vybe-test: csharp/csharp_pattern_matching/switch_expression_with_null_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string value = null;
string result = value switch { null => "nothing", var s => s };
__Check((result).ToString(), "nothing");

// vybe-test: csharp/csharp_switch_type_patterns/is_pattern_with_null_constant_detects_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text = null;
__Check((text is null).ToString(), "True");

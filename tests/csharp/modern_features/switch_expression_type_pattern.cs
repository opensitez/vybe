// vybe-test: csharp/modern_features/switch_expression_type_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object obj = 42;
string result = obj switch {
    int i => "int: " + i,
    string s => "string: " + s,
    _ => "unknown"
};
__Check((result).ToString(), "int: 42");

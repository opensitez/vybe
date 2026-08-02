// vybe-test: csharp/csharp_switch_type_patterns/switch_expression_returns_value_from_matching_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 2;
string word = n switch { 1 => "one", 2 => "two", _ => "many" };
__Check((word).ToString(), "two");

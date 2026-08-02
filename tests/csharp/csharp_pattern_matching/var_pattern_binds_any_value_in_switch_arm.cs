// vybe-test: csharp/csharp_pattern_matching/var_pattern_binds_any_value_in_switch_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o = 42;
string result = o switch { var x when x is int n && n > 10 => "big int", _ => "other" };
__Check((result).ToString(), "big int");

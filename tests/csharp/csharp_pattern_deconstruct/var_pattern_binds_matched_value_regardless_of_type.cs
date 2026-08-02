// vybe-test: csharp/csharp_pattern_deconstruct/var_pattern_binds_matched_value_regardless_of_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object value = 42;
if (value is var captured) __Check((captured).ToString(), "42");

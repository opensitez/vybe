// vybe-test: csharp/csharp_pattern_matching_advanced/var_pattern_binds_any_value_in_switch_statement
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object item = 18; switch (item) { case var anything: __Check((anything).ToString(), "18"); break; }

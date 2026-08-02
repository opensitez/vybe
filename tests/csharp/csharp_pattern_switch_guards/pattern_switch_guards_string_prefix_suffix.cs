// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_switch_guards
string feature = "pattern_switch_guards"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");

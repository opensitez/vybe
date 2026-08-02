// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_switch_guards
string feature = "pattern_switch_guards:42"; __Check((feature.Length >= 1).ToString(), "True");

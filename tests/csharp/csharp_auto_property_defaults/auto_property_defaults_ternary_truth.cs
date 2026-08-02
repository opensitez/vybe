// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
int seed = 65; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");

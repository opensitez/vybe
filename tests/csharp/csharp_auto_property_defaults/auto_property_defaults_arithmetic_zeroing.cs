// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
int seed = 65; __Check((seed - seed == 0).ToString(), "True");

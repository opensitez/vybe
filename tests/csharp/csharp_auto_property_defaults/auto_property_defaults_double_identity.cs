// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
double seed = 65; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");

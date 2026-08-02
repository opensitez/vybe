// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
string feature = "auto_property_defaults"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");

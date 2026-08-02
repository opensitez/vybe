// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");

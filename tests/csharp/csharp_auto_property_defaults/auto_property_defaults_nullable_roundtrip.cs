// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
int? maybe = 65; __Check((maybe.HasValue && maybe.Value == 65).ToString(), "True");

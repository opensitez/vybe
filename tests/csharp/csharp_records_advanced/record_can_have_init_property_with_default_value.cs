// vybe-test: csharp/csharp_records_advanced/record_can_have_init_property_with_default_value
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record User { public string Name { get; init; } = "guest"; } __Check((new User().Name).ToString(), "guest");

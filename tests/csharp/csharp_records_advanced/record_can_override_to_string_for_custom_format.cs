// vybe-test: csharp/csharp_records_advanced/record_can_override_to_string_for_custom_format
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record User(string Name) { public override string ToString() { return $"User:{Name}"; } } __Check((new User("Ada")).ToString(), "User:Ada");

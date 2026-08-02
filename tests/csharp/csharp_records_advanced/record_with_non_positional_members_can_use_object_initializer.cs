// vybe-test: csharp/csharp_records_advanced/record_with_non_positional_members_can_use_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Theme { public string Name { get; init; } public int Version { get; init; } } var theme = new Theme { Name = "light", Version = 2 }; __Check((theme.Name + ":" + theme.Version).ToString(), "light:2");

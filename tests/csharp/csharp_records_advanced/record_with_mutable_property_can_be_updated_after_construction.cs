// vybe-test: csharp/csharp_records_advanced/record_with_mutable_property_can_be_updated_after_construction
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Box { public int Value { get; set; } } var box = new Box { Value = 3 }; box.Value = 8; __Check((box.Value).ToString(), "8");

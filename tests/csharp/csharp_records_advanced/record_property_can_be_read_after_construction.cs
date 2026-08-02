// vybe-test: csharp/csharp_records_advanced/record_property_can_be_read_after_construction
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Config(string Name); var config = new Config("debug"); __Check((config.Name).ToString(), "debug");

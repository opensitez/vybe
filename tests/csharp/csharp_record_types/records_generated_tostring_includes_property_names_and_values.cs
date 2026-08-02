// vybe-test: csharp/csharp_record_types/records_generated_tostring_includes_property_names_and_values
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Tag(string Name);
__Check((new Tag("admin").ToString().Contains("admin")).ToString(), "True");

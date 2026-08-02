// vybe-test: csharp/csharp_record_struct/record_struct_to_string_includes_property_values
// origin: languages/csharp/tests/csharp/test_csharp_record_struct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Tag(string Name);
__Check((new Tag("admin").ToString().Contains("admin")).ToString(), "True");

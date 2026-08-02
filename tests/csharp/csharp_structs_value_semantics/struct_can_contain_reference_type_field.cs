// vybe-test: csharp/csharp_structs_value_semantics/struct_can_contain_reference_type_field
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Wrapper { public string Name; } var wrapper = new Wrapper { Name = "text" }; __Check((wrapper.Name).ToString(), "text");

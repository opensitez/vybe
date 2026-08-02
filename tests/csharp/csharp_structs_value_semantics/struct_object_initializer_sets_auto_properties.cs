// vybe-test: csharp/csharp_structs_value_semantics/struct_object_initializer_sets_auto_properties
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Box { public int Value { get; set; } } var box = new Box { Value = 11 }; __Check((box.Value).ToString(), "11");

// vybe-test: csharp/csharp_structs_value_semantics/nested_struct_inside_class_is_constructible
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer { public struct Inner { public int Value; } } var value = new Outer.Inner { Value = 8 }; __Check((value.Value).ToString(), "8");

// vybe-test: csharp/csharp_structs_value_semantics/struct_equals_compares_field_values
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point { public int X; public int Y; } var left = new Point { X = 1, Y = 2 }; var right = new Point { X = 1, Y = 2 }; __Check((left.Equals(right)).ToString(), "True");

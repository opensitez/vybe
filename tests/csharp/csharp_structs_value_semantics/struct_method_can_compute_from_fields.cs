// vybe-test: csharp/csharp_structs_value_semantics/struct_method_can_compute_from_fields
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point { public int X; public int Y; public int Sum() { return X + Y; } } var point = new Point { X = 4, Y = 6 }; __Check((point.Sum()).ToString(), "10");

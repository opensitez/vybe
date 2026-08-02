// vybe-test: csharp/csharp_structs_value_semantics/default_struct_fields_start_at_zero
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point { public int X; public int Y; } var point = new Point(); __Check((point.X).ToString(), "0"); __Check((point.Y).ToString(), "0");

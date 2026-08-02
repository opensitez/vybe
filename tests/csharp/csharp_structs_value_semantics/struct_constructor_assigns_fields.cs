// vybe-test: csharp/csharp_structs_value_semantics/struct_constructor_assigns_fields
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Point { public int X; public int Y; public Point(int x, int y) { X = x; Y = y; } } var point = new Point(2, 3); __Check((point.X + point.Y).ToString(), "5");

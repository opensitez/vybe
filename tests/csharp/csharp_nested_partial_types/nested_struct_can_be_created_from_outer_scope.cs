// vybe-test: csharp/csharp_nested_partial_types/nested_struct_can_be_created_from_outer_scope
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Geometry {
    public struct Point {
        public int X;
        public int Y;
    }
}
var point = new Geometry.Point { X = 3, Y = 4 };
__Check((point.X + point.Y).ToString(), "7");

// vybe-test: csharp/csharp_collection_initializer_syntax/object_initializer_sets_public_fields_before_use
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point { public int X; public int Y; }
var point = new Point { X = 2, Y = 5 };
__Check((point.X + point.Y).ToString(), "7");

// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_struct_field_access
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Shapes;
struct Point { public int X; public int Y; }
var p = new Point { X = 2, Y = 3 };
__Check((p.X + p.Y).ToString(), "5");

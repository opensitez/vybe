// vybe-test: csharp/csharp_namespace_aliases/namespace_can_contain_nested_struct_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo { public struct Point { public int X; public int Y; } } var point = new Demo.Point { X = 2, Y = 5 }; __Check((point.X + point.Y).ToString(), "7");

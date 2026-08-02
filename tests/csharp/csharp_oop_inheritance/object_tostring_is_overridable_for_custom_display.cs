// vybe-test: csharp/csharp_oop_inheritance/object_tostring_is_overridable_for_custom_display
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point { public int X,Y; public override string ToString() => $"({X},{Y})"; }
__Check((new Point { X=1, Y=2 }).ToString(), "(1,2)");

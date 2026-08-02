// vybe-test: csharp/csharp_oop_inheritance/constructor_chaining_with_base_passes_args_to_parent
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shape { public string Color; public Shape(string c) { Color = c; } }
class Box : Shape { public Box(string c) : base(c) { } }
__Check((new Box("red").Color).ToString(), "red");

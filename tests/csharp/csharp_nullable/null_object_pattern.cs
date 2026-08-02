// vybe-test: csharp/csharp_nullable/null_object_pattern
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box {
    public int Value;
    public Box(int v) { Value = v; }
}
Box b = null;
__Check((b == null).ToString(), "True");
b = new Box(42);
__Check((b == null).ToString(), "False");
__Check((b.Value).ToString(), "42");

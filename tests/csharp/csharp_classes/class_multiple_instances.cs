// vybe-test: csharp/csharp_classes/class_multiple_instances
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

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
var a = new Box(10);
var b = new Box(20);
__Check((a.Value + b.Value).ToString(), "30");

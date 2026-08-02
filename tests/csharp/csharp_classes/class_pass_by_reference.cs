// vybe-test: csharp/csharp_classes/class_pass_by_reference
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box {
    public int Value;
}
void Modify(Box b) {
    b.Value = 99;
}
var b = new Box();
b.Value = 1;
Modify(b);
__Check((b.Value).ToString(), "99");

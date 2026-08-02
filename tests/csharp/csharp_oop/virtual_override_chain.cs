// vybe-test: csharp/csharp_oop/virtual_override_chain
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A {
    public virtual string Name() { return "A"; }
}
class B : A {
    public override string Name() { return "B"; }
}
var obj = new B();
__Check((obj.Name()).ToString(), "B");

// vybe-test: csharp/csharp_classes/multi_level_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A {
    public virtual string Who() { return "A"; }
}
class B : A {
    public override string Who() { return "B->" + base.Who(); }
}
class C : B {
    public override string Who() { return "C->" + base.Who(); }
}
var c = new C();
__Check((c.Who()).ToString(), "C->B->A");

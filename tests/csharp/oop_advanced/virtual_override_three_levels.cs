// vybe-test: csharp/oop_advanced/virtual_override_three_levels
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

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
    public override string Who() { return "B"; }
}
class C : B {
    public override string Who() { return "C"; }
}
A obj = new C();
__Check((obj.Who()).ToString(), "C");

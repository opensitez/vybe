// vybe-test: csharp/oop_advanced/virtual_override_three_levels
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

A obj = new C();
__P((obj.Who()).ToString());
__Check("C");

class A {
    public virtual string Who() { return "A"; }
}

class B : A {
    public override string Who() { return "B"; }
}

class C : B {
    public override string Who() { return "C"; }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}

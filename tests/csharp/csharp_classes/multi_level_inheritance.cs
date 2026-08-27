// vybe-test: csharp/csharp_classes/multi_level_inheritance
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var c = new C();
__P((c.Who()).ToString());
__Check("C->B->A");

class A {
    public virtual string Who() { return "A"; }
}

class B : A {
    public override string Who() { return "B->" + base.Who(); }
}

class C : B {
    public override string Who() { return "C->" + base.Who(); }
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

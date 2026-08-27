// vybe-test: csharp/csharp_oop/virtual_override_chain
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var obj = new B();
__P((obj.Name()).ToString());
__Check("B");

class A {
    public virtual string Name() { return "A"; }
}

class B : A {
    public override string Name() { return "B"; }
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

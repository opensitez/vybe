// vybe-test: csharp/csharp_virtual_dispatch_semantics/base_keyword_invokes_parent_implementation_from_override
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

__P((new StepCounter().Next()).ToString());
__Check("3");

class Counter {
    public virtual int Next() { return 1; }
}

class StepCounter : Counter {
    public override int Next() { return base.Next() + 2; }
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

// vybe-test: csharp/csharp_abstract_class/abstract_class_holding_state_shared_with_subclass
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

using static __Harness;

var c=new By2();
c.Increment();
c.Increment();
__P((c.Value).ToString());
__Check("4");

abstract class Counter{
    protected int Count;
    public abstract void Increment();
    public int Value=>Count;
}

class By2:Counter{public override void Increment(){Count+=2;}}

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

// vybe-test: csharp/csharp_oop_advanced2/abstract_class_partial_implementation_forces_concrete_override
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

using static __Harness;

__P((new Alpha().Run()).ToString());
__Check("run:alpha");

abstract class Step{
    public abstract string Name();
    public string Run()=>$"run:{Name()}";
}

class Alpha:Step{public override string Name()=>"alpha";}

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

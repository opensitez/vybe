// vybe-test: csharp/csharp_oop_advanced2/covariant_return_type_narrows_return_of_override
// origin: languages/csharp/tests/csharp/test_csharp_oop_advanced2.rs

using static __Harness;

Derived d=new Derived();
__P((d.Create()).ToString());
__Check("derived");

class Base{public virtual object Create()=>new object();}

class Derived:Base{public override string Create()=>"derived";}

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

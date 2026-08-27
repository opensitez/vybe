// vybe-test: csharp/csharp_nested_classes/nested_class_can_implement_interface
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

using static __Harness;

IValue v=new Host.Impl();
__P((v.Get()).ToString());
__Check("5");

interface IValue{int Get();}

class Host{public class Impl:IValue{public int Get()=>5;}}

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
